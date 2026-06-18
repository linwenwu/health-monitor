import React, { useState, useMemo, useRef, useCallback, useEffect } from "react";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer, ReferenceLine, ReferenceArea,
  BarChart, Bar, Cell, AreaChart, Area, ComposedChart
} from "recharts";
import * as XLSX from "xlsx";

// ============================================================
// CONSTANTS
// ============================================================

const COLORS = {
  primary: "#3b82f6",
  success: "#10b981",
  warning: "#f59e0b",
  danger: "#ef4444",
  purple: "#8b5cf6",
  sidebarBg: "#0f172a",
  sidebarHover: "#1e293b",
  sidebarActive: "#2563eb",
  cardBg: "#ffffff",
  pageBg: "#f1f5f9",
  textDark: "#1e293b",
  textMid: "#64748b",
  textLight: "#94a3b8",
  border: "#e2e8f0",
  glucoseLine: "#f97316",
  bpSystolic: "#ef4444",
  bpDiastolic: "#3b82f6",
  pulse: "#8b5cf6",
 达标: "#10b981",
  偏高: "#f59e0b",
  超标: "#ef4444",
};

const MEAL_TIMES = [
  { key: "fasting", label: "空腹" },
  { key: "pre_breakfast", label: "早餐前" },
  { key: "post_breakfast", label: "早餐后2h" },
  { key: "pre_lunch", label: "午餐前" },
  { key: "post_lunch", label: "午餐后2h" },
  { key: "pre_dinner", label: "晚餐前" },
  { key: "post_dinner", label: "晚餐后2h" },
  { key: "bedtime", label: "睡前" },
  { key: "random", label: "随机" },
];

const TARGETS = {
  glucose: { min: 3.9, max: 10.0, fastingMax: 7.0, postMealMax: 10.0, unit: "mmol/L" },
  bpSystolic: { min: 90, max: 140, unit: "mmHg" },
  bpDiastolic: { min: 60, max: 90, unit: "mmHg" },
  pulse: { min: 60, max: 100, unit: "bpm" },
};

const TABS = [
  { id: "dashboard", label: "数据面板", icon: "📊" },
  { id: "glucose", label: "血糖录入", icon: "🩸" },
  { id: "bp", label: "血压录入", icon: "💓" },
  { id: "import", label: "文件导入", icon: "📁" },
  { id: "history", label: "历史记录", icon: "📋" },
  { id: "analysis", label: "深度分析", icon: "🔍" },
];

// ============================================================
// SAMPLE DATA (from 手动录入的数据.xlsx)
// ============================================================

const SAMPLE_GLUCOSE = [
  { id: "g1", date: "2026-05-28", time: "07:15", value: 9.2, mealTime: "fasting", note: "确诊日" },
  { id: "g2", date: "2026-05-28", time: "22:00", value: 7.8, mealTime: "bedtime", note: "" },
  { id: "g3", date: "2026-05-29", time: "07:00", value: 7.6, mealTime: "fasting", note: "" },
  { id: "g4", date: "2026-05-29", time: "21:30", value: 8.2, mealTime: "bedtime", note: "" },
  { id: "g5", date: "2026-05-30", time: "07:10", value: 6.8, mealTime: "fasting", note: "" },
  { id: "g6", date: "2026-05-30", time: "16:00", value: 9.1, mealTime: "random", note: "" },
  { id: "g7", date: "2026-05-31", time: "07:00", value: 7.3, mealTime: "fasting", note: "" },
  { id: "g8", date: "2026-05-31", time: "22:00", value: 8.9, mealTime: "bedtime", note: "" },
  { id: "g9", date: "2026-06-01", time: "07:15", value: 8.1, mealTime: "fasting", note: "" },
  { id: "g10", date: "2026-06-01", time: "12:30", value: 10.2, mealTime: "post_lunch", note: "" },
  { id: "g11", date: "2026-06-01", time: "22:00", value: 9.5, mealTime: "bedtime", note: "" },
  { id: "g12", date: "2026-06-02", time: "07:00", value: 8.4, mealTime: "fasting", note: "" },
  { id: "g13", date: "2026-06-02", time: "19:30", value: 8.0, mealTime: "post_dinner", note: "" },
  { id: "g14", date: "2026-06-03", time: "07:10", value: 7.9, mealTime: "fasting", note: "" },
  { id: "g15", date: "2026-06-03", time: "17:00", value: 8.9, mealTime: "random", note: "偷吃鹌鹑蛋" },
  { id: "g16", date: "2026-06-03", time: "21:30", value: 8.7, mealTime: "bedtime", note: "" },
  { id: "g17", date: "2026-06-04", time: "07:00", value: 6.5, mealTime: "fasting", note: "首次空腹达标！" },
  { id: "g18", date: "2026-06-04", time: "12:00", value: 9.8, mealTime: "post_lunch", note: "" },
  { id: "g19", date: "2026-06-04", time: "22:00", value: 8.4, mealTime: "bedtime", note: "" },
  { id: "g20", date: "2026-06-05", time: "07:10", value: 7.2, mealTime: "fasting", note: "" },
  { id: "g21", date: "2026-06-05", time: "19:30", value: 7.5, mealTime: "post_dinner", note: "" },
  { id: "g22", date: "2026-06-06", time: "07:00", value: 6.7, mealTime: "fasting", note: "" },
  { id: "g23", date: "2026-06-06", time: "22:00", value: 8.1, mealTime: "bedtime", note: "" },
  { id: "g24", date: "2026-06-07", time: "07:10", value: 6.2, mealTime: "fasting", note: "" },
  { id: "g25", date: "2026-06-07", time: "16:00", value: 11.7, mealTime: "random", note: "偷吃饼干" },
  { id: "g26", date: "2026-06-08", time: "07:00", value: 6.9, mealTime: "fasting", note: "" },
  { id: "g27", date: "2026-06-08", time: "22:00", value: 7.8, mealTime: "bedtime", note: "" },
  { id: "g28", date: "2026-06-09", time: "07:10", value: 6.3, mealTime: "fasting", note: "连续六天达标" },
  { id: "g29", date: "2026-06-09", time: "21:30", value: 7.5, mealTime: "bedtime", note: "" },
  { id: "g30", date: "2026-06-10", time: "07:00", value: 6.4, mealTime: "fasting", note: "停药日（伊伐布雷定+美托洛尔）" },
  { id: "g31", date: "2026-06-10", time: "12:30", value: 9.0, mealTime: "post_lunch", note: "" },
  { id: "g32", date: "2026-06-11", time: "07:10", value: 6.8, mealTime: "fasting", note: "" },
  { id: "g33", date: "2026-06-11", time: "22:00", value: 7.9, mealTime: "bedtime", note: "" },
  { id: "g34", date: "2026-06-12", time: "07:00", value: 7.1, mealTime: "fasting", note: "" },
  { id: "g35", date: "2026-06-12", time: "19:30", value: 7.6, mealTime: "post_dinner", note: "" },
  { id: "g36", date: "2026-06-13", time: "07:10", value: 6.5, mealTime: "fasting", note: "" },
  { id: "g37", date: "2026-06-13", time: "22:00", value: 7.2, mealTime: "bedtime", note: "" },
  { id: "g38", date: "2026-06-14", time: "07:00", value: 5.6, mealTime: "fasting", note: "最低值！" },
  { id: "g39", date: "2026-06-14", time: "12:00", value: 8.3, mealTime: "post_lunch", note: "" },
  { id: "g40", date: "2026-06-15", time: "07:10", value: 6.7, mealTime: "fasting", note: "" },
  { id: "g41", date: "2026-06-15", time: "21:30", value: 7.8, mealTime: "bedtime", note: "" },
  { id: "g42", date: "2026-06-16", time: "07:00", value: 6.9, mealTime: "fasting", note: "" },
  { id: "g43", date: "2026-06-16", time: "19:30", value: 8.0, mealTime: "post_dinner", note: "" },
  { id: "g44", date: "2026-06-17", time: "07:10", value: 6.4, mealTime: "fasting", note: "" },
  { id: "g45", date: "2026-06-17", time: "22:00", value: 7.5, mealTime: "bedtime", note: "" },
  { id: "g46", date: "2026-06-18", time: "07:00", value: 6.6, mealTime: "fasting", note: "" },
  { id: "g47", date: "2026-06-18", time: "12:30", value: 8.5, mealTime: "post_lunch", note: "" },
];

const SAMPLE_BP = [
  { id: "b1", date: "2026-05-28", time: "07:20", systolic: 112, diastolic: 72, pulse: 63 },
  { id: "b2", date: "2026-05-29", time: "07:15", systolic: 108, diastolic: 68, pulse: 60 },
  { id: "b3", date: "2026-05-30", time: "07:20", systolic: 105, diastolic: 65, pulse: 58 },
  { id: "b4", date: "2026-05-31", time: "07:15", systolic: 110, diastolic: 70, pulse: 62 },
  { id: "b5", date: "2026-06-01", time: "07:20", systolic: 95, diastolic: 60, pulse: 55 },
  { id: "b6", date: "2026-06-02", time: "07:15", systolic: 98, diastolic: 62, pulse: 57 },
  { id: "b7", date: "2026-06-03", time: "07:20", systolic: 102, diastolic: 65, pulse: 60 },
  { id: "b8", date: "2026-06-04", time: "07:15", systolic: 100, diastolic: 63, pulse: 58 },
  { id: "b9", date: "2026-06-05", time: "07:20", systolic: 96, diastolic: 61, pulse: 56 },
  { id: "b10", date: "2026-06-06", time: "07:15", systolic: 99, diastolic: 64, pulse: 59 },
  { id: "b11", date: "2026-06-07", time: "07:20", systolic: 103, diastolic: 66, pulse: 61 },
  { id: "b12", date: "2026-06-08", time: "07:15", systolic: 97, diastolic: 62, pulse: 57 },
  { id: "b13", date: "2026-06-09", time: "07:20", systolic: 101, diastolic: 64, pulse: 60 },
  { id: "b14", date: "2026-06-10", time: "07:15", systolic: 105, diastolic: 68, pulse: 65, note: "停药日" },
  { id: "b15", date: "2026-06-11", time: "07:20", systolic: 108, diastolic: 70, pulse: 68 },
  { id: "b16", date: "2026-06-12", time: "07:15", systolic: 112, diastolic: 72, pulse: 70 },
  { id: "b17", date: "2026-06-13", time: "07:20", systolic: 106, diastolic: 68, pulse: 66 },
  { id: "b18", date: "2026-06-14", time: "07:15", systolic: 115, diastolic: 75, pulse: 72 },
  { id: "b19", date: "2026-06-15", time: "07:20", systolic: 110, diastolic: 70, pulse: 68 },
  { id: "b20", date: "2026-06-16", time: "07:15", systolic: 108, diastolic: 69, pulse: 65 },
  { id: "b21", date: "2026-06-17", time: "08:18", systolic: 118, diastolic: 74, pulse: 70 },
  { id: "b22", date: "2026-06-18", time: "08:30", systolic: 124, diastolic: 80, pulse: 75 },
];

// Medication events
const EVENTS = [
  { date: "2026-06-10", label: "停用伊伐布雷定+美托洛尔", color: COLORS.danger },
];

// ============================================================
// HELPER FUNCTIONS
// ============================================================

const fmtDate = (d) => {
  if (!d) return "";
  const parts = d.split("-");
  return `${parts[1]}/${parts[2]}`;
};

const getMealLabel = (key) => {
  const found = MEAL_TIMES.find((m) => m.key === key);
  return found ? found.label : key || "";
};

const classifyGlucose = (value, mealTime) => {
  if (value < 3.9) return "偏低";
  if (mealTime === "fasting" || mealTime === "pre_breakfast" || mealTime === "pre_lunch" || mealTime === "pre_dinner") {
    return value <= TARGETS.glucose.fastingMax ? "达标" : "超标";
  }
  return value <= TARGETS.glucose.postMealMax ? "达标" : "超标";
};

const classifyBP = (sys, dia) => {
  if (sys >= TARGETS.bpSystolic.min && sys <= TARGETS.bpSystolic.max &&
      dia >= TARGETS.bpDiastolic.min && dia <= TARGETS.bpDiastolic.max) {
    return "达标";
  }
  if (sys > TARGETS.bpSystolic.max || dia > TARGETS.bpDiastolic.max) return "偏高";
  return "偏低";
};

const calcStats = (arr, key) => {
  if (!arr.length) return { avg: 0, min: 0, max: 0 };
  const vals = arr.map((r) => r[key]).filter((v) => v != null);
  if (!vals.length) return { avg: 0, min: 0, max: 0 };
  const sum = vals.reduce((a, b) => a + b, 0);
  return {
    avg: (sum / vals.length).toFixed(1),
    min: Math.min(...vals),
    max: Math.max(...vals),
  };
};

const calc达标率 = (arr, classifyFn) => {
  if (!arr.length) return 0;
  const 达标 = arr.filter((r) => classifyFn(r) === "达标").length;
  return Math.round((达标 / arr.length) * 100);
};

const genId = () => `id_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;

// ============================================================
// SUB-COMPONENTS
// ============================================================

function StatCard({ title, value, unit, icon, color, trend, subtext }) {
  return (
    <div style={{
      background: COLORS.cardBg,
      borderRadius: 12,
      padding: "20px 24px",
      boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
      border: `1px solid ${COLORS.border}`,
      minWidth: 0,
      flex: 1,
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <span style={{ fontSize: 13, color: COLORS.textMid, fontWeight: 500 }}>{title}</span>
        <span style={{ fontSize: 22 }}>{icon}</span>
      </div>
      <div style={{ display: "flex", alignItems: "baseline", gap: 6 }}>
        <span style={{ fontSize: 28, fontWeight: 700, color: color || COLORS.textDark }}>{value}</span>
        <span style={{ fontSize: 13, color: COLORS.textLight }}>{unit}</span>
      </div>
      {subtext && (
        <div style={{ fontSize: 12, color: COLORS.textMid, marginTop: 6 }}>{subtext}</div>
      )}
      {trend && (
        <div style={{ fontSize: 12, color: trend > 0 ? COLORS.danger : COLORS.success, marginTop: 4 }}>
          {trend > 0 ? "↑" : "↓"} 较前日 {Math.abs(trend).toFixed(1)}
        </div>
      )}
    </div>
  );
}

function SectionCard({ title, children, extra }) {
  return (
    <div style={{
      background: COLORS.cardBg,
      borderRadius: 12,
      padding: 24,
      boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
      border: `1px solid ${COLORS.border}`,
      marginBottom: 20,
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <h3 style={{ margin: 0, fontSize: 16, fontWeight: 600, color: COLORS.textDark }}>{title}</h3>
        {extra}
      </div>
      {children}
    </div>
  );
}

function Toast({ message, type, onClose }) {
  useEffect(() => {
    const timer = setTimeout(onClose, 3000);
    return () => clearTimeout(timer);
  }, [onClose]);

  const bg = type === "success" ? COLORS.success : type === "error" ? COLORS.danger : COLORS.primary;
  return (
    <div style={{
      position: "fixed", top: 24, right: 24, zIndex: 9999,
      background: bg, color: "#fff", padding: "12px 20px",
      borderRadius: 8, fontSize: 14, fontWeight: 500,
      boxShadow: "0 4px 12px rgba(0,0,0,0.2)", cursor: "pointer",
      animation: "fadeIn 0.2s ease",
    }} onClick={onClose}>
      {type === "success" ? "✓ " : type === "error" ? "✗ " : "ℹ "}{message}
    </div>
  );
}

function EmptyState({ icon, title, desc }) {
  return (
    <div style={{ textAlign: "center", padding: "60px 20px", color: COLORS.textMid }}>
      <div style={{ fontSize: 48, marginBottom: 12 }}>{icon}</div>
      <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>{title}</div>
      <div style={{ fontSize: 13 }}>{desc}</div>
    </div>
  );
}

// ============================================================
// DASHBOARD VIEW
// ============================================================

function DashboardView({ glucose, bp }) {
  const glucoseStats = calcStats(glucose, "value");
  const bpStats_sys = calcStats(bp, "systolic");
  const bpStats_dia = calcStats(bp, "diastolic");
  const pulseStats = calcStats(bp, "pulse");

  const 血糖达标率 = calc达标率(glucose, (r) => classifyGlucose(r.value, r.mealTime));
  const bp达标率 = calc达标率(bp, (r) => classifyBP(r.systolic, r.diastolic));

  // Prepare chart data - glucose
  const glucoseChartData = useMemo(() => {
    const byDate = {};
    glucose.forEach((r) => {
      if (!byDate[r.date]) byDate[r.date] = { date: r.date, values: [], readings: [] };
      byDate[r.date].values.push(r.value);
      byDate[r.date].readings.push(r);
    });
    return Object.values(byDate)
      .sort((a, b) => a.date.localeCompare(b.date))
      .map((d) => ({
        date: fmtDate(d.date),
        fullDate: d.date,
        avg: +(d.values.reduce((a, b) => a + b, 0) / d.values.length).toFixed(1),
        min: Math.min(...d.values),
        max: Math.max(...d.values),
        count: d.values.length,
        isEvent: EVENTS.some((e) => e.date === d.fullDate),
      }));
  }, [glucose]);

  // Prepare chart data - BP
  const bpChartData = useMemo(() => {
    return [...bp]
      .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))
      .map((r) => ({
        date: fmtDate(r.date),
        fullDate: r.date,
        systolic: r.systolic,
        diastolic: r.diastolic,
        pulse: r.pulse,
        isEvent: EVENTS.some((e) => e.date === r.date),
      }));
  }, [bp]);

  // Glucose trend day-over-day
  const lastTwoDays = glucoseChartData.slice(-2);
  const glucoseTrend = lastTwoDays.length === 2 ? lastTwoDays[1].avg - lastTwoDays[0].avg : 0;

  // Distribution by meal time
  const mealDistribution = useMemo(() => {
    const groups = {};
    MEAL_TIMES.forEach((m) => { groups[m.key] = []; });
    glucose.forEach((r) => {
      const key = r.mealTime || "random";
      if (!groups[key]) groups[key] = [];
      groups[key].push(r.value);
    });
    return MEAL_TIMES
      .filter((m) => groups[m.key] && groups[m.key].length > 0)
      .map((m) => ({
        name: m.label,
        avg: +(groups[m.key].reduce((a, b) => a + b, 0) / groups[m.key].length).toFixed(1),
        count: groups[m.key].length,
        达标: groups[m.key].filter((v) => {
          if (m.key === "fasting" || m.key === "pre_breakfast" || m.key === "pre_lunch" || m.key === "pre_dinner") {
            return v <= TARGETS.glucose.fastingMax;
          }
          return v <= TARGETS.glucose.postMealMax;
        }).length,
      }));
  }, [glucose]);

  const CustomTooltip = ({ active, payload }) => {
    if (!active || !payload || !payload.length) return null;
    const data = payload[0].payload;
    return (
      <div style={{
        background: "#fff", border: `1px solid ${COLORS.border}`,
        borderRadius: 8, padding: "10px 14px", fontSize: 13,
        boxShadow: "0 2px 8px rgba(0,0,0,0.12)",
      }}>
        <div style={{ fontWeight: 600, marginBottom: 4 }}>{data.fullDate || data.date}</div>
        {payload.map((p, i) => (
          <div key={i} style={{ color: p.color, marginBottom: 2 }}>
            {p.name}: <strong>{p.value}</strong>
          </div>
        ))}
        {data.isEvent && <div style={{ color: COLORS.danger, fontWeight: 600, marginTop: 4 }}>⚠ 停药日</div>}
      </div>
    );
  };

  return (
    <div>
      <h2 style={{ margin: "0 0 20px", fontSize: 20, fontWeight: 700, color: COLORS.textDark }}>
        健康监测面板
      </h2>

      {/* Stat Cards */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16, marginBottom: 24 }}>
        <StatCard
          title="平均血糖" value={glucoseStats.avg} unit="mmol/L"
          icon="🩸" color={COLORS.glucoseLine}
          subtext={`范围 ${glucoseStats.min} ~ ${glucoseStats.max}`}
          trend={glucoseTrend}
        />
        <StatCard
          title="血糖达标率" value={`${血糖达标率}%`}
          icon="✅" color={血糖达标率 >= 70 ? COLORS.success : COLORS.warning}
          subtext={`共 ${glucose.length} 条记录`}
        />
        <StatCard
          title="平均血压" value={`${bpStats_sys.avg}/${bpStats_dia.avg}`} unit="mmHg"
          icon="💓" color={COLORS.bpSystolic}
          subtext={`脉搏 ${pulseStats.avg} bpm`}
        />
        <StatCard
          title="血压达标率" value={`${bp达标率}%`}
          icon="❤️" color={bp达标率 >= 80 ? COLORS.success : COLORS.warning}
          subtext={`共 ${bp.length} 条记录`}
        />
      </div>

      {/* Glucose Trend Chart */}
      <SectionCard title="📈 血糖趋势（日均值）" extra={
        <span style={{ fontSize: 12, color: COLORS.textMid }}>
          目标：空腹 &lt;{TARGETS.glucose.fastingMax} / 餐后 &lt;{TARGETS.glucose.postMealMax} mmol/L
        </span>
      }>
        <ResponsiveContainer width="100%" height={280}>
          <ComposedChart data={glucoseChartData} margin={{ top: 10, right: 20, bottom: 0, left: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
            <XAxis dataKey="date" tick={{ fontSize: 12, fill: COLORS.textMid }} />
            <YAxis domain={[3, 13]} tick={{ fontSize: 12, fill: COLORS.textMid }} />
            <Tooltip content={<CustomTooltip />} />
            <ReferenceArea
              y1={TARGETS.glucose.min} y2={TARGETS.glucose.postMealMax}
              fill="#10b98120" stroke="none"
            />
            <ReferenceLine
              y={TARGETS.glucose.fastingMax} stroke={COLORS.success}
              strokeDasharray="5 5" label={{ value: "空腹目标", position: "right", fontSize: 11 }}
            />
            <ReferenceLine
              y={TARGETS.glucose.postMealMax} stroke={COLORS.warning}
              strokeDasharray="5 5" label={{ value: "餐后目标", position: "right", fontSize: 11 }}
            />
            <Area type="monotone" dataKey="max" fill="#f9731620" stroke="none" name="最高" />
            <Area type="monotone" dataKey="min" fill="#fff" stroke="none" name="最低" />
            <Line
              type="monotone" dataKey="avg" stroke={COLORS.glucoseLine}
              strokeWidth={2.5} dot={{ r: 4, fill: COLORS.glucoseLine, strokeWidth: 0 }}
              activeDot={{ r: 6 }} name="日均值"
            />
            {EVENTS.map((e, i) => (
              <ReferenceLine
                key={i} x={fmtDate(e.date)} stroke={e.color}
                strokeDasharray="3 3" strokeWidth={2}
                label={{ value: e.label, position: "top", fontSize: 11, fill: e.color }}
              />
            ))}
          </ComposedChart>
        </ResponsiveContainer>
      </SectionCard>

      {/* BP Trend Chart */}
      <SectionCard title="📉 血压脉搏趋势" extra={
        <span style={{ fontSize: 12, color: COLORS.textMid }}>
          正常：{TARGETS.bpSystolic.min}-{TARGETS.bpSystolic.max}/{TARGETS.bpDiastolic.min}-{TARGETS.bpDiastolic.max} mmHg
        </span>
      }>
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={bpChartData} margin={{ top: 10, right: 20, bottom: 0, left: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
            <XAxis dataKey="date" tick={{ fontSize: 12, fill: COLORS.textMid }} />
            <YAxis yAxisId="bp" domain={[40, 160]} tick={{ fontSize: 12, fill: COLORS.textMid }} />
            <YAxis yAxisId="pulse" orientation="right" domain={[40, 110]} tick={{ fontSize: 12, fill: COLORS.pulse }} />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ fontSize: 13 }} />
            <ReferenceArea yAxisId="bp" y1={TARGETS.bpSystolic.min} y2={TARGETS.bpSystolic.max} fill="#3b82f610" stroke="none" />
            <Line
              yAxisId="bp" type="monotone" dataKey="systolic" name="高压"
              stroke={COLORS.bpSystolic} strokeWidth={2}
              dot={{ r: 3, fill: COLORS.bpSystolic, strokeWidth: 0 }}
            />
            <Line
              yAxisId="bp" type="monotone" dataKey="diastolic" name="低压"
              stroke={COLORS.bpDiastolic} strokeWidth={2}
              dot={{ r: 3, fill: COLORS.bpDiastolic, strokeWidth: 0 }}
            />
            <Line
              yAxisId="pulse" type="monotone" dataKey="pulse" name="脉搏"
              stroke={COLORS.pulse} strokeWidth={2} strokeDasharray="5 5"
              dot={{ r: 3, fill: COLORS.pulse, strokeWidth: 0 }}
            />
            {EVENTS.map((e, i) => (
              <ReferenceLine
                key={i} x={fmtDate(e.date)} yAxisId="bp" stroke={e.color}
                strokeDasharray="3 3" strokeWidth={2}
              />
            ))}
          </LineChart>
        </ResponsiveContainer>
      </SectionCard>

      {/* Meal Time Distribution */}
      <SectionCard title="📊 各时段平均血糖">
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={mealDistribution} margin={{ top: 10, right: 20, bottom: 0, left: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
            <XAxis dataKey="name" tick={{ fontSize: 12, fill: COLORS.textMid }} />
            <YAxis domain={[0, 12]} tick={{ fontSize: 12, fill: COLORS.textMid }} />
            <Tooltip
              formatter={(val, name) => name === "avg" ? [`${val} mmol/L`, "平均值"] : [val, name]}
              contentStyle={{ fontSize: 13, borderRadius: 8 }}
            />
            <ReferenceLine y={TARGETS.glucose.fastingMax} stroke={COLORS.success} strokeDasharray="5 5" />
            <Bar dataKey="avg" name="平均值" radius={[6, 6, 0, 0]}>
              {mealDistribution.map((entry, index) => (
                <Cell
                  key={index}
                  fill={entry.avg <= TARGETS.glucose.fastingMax ? COLORS.success :
                    entry.avg <= TARGETS.glucose.postMealMax ? COLORS.warning : COLORS.danger}
                />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </SectionCard>
    </div>
  );
}

// ============================================================
// GLUCOSE ENTRY VIEW
// ============================================================

function GlucoseEntryView({ onSave, showToast }) {
  const [form, setForm] = useState({
    date: new Date().toISOString().split("T")[0],
    time: new Date().toTimeString().slice(0, 5),
    value: "",
    mealTime: "fasting",
    note: "",
  });

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!form.value || isNaN(parseFloat(form.value))) {
      showToast("请输入有效的血糖值", "error");
      return;
    }
    onSave({
      id: genId(),
      date: form.date,
      time: form.time,
      value: parseFloat(form.value),
      mealTime: form.mealTime,
      note: form.note,
    });
    showToast("血糖记录已保存", "success");
    setForm({ ...form, value: "", note: "" });
  };

  const inputStyle = {
    width: "100%", padding: "10px 14px", border: `1px solid ${COLORS.border}`,
    borderRadius: 8, fontSize: 14, outline: "none", boxSizing: "border-box",
    transition: "border-color 0.2s",
  };

  const labelStyle = { display: "block", fontSize: 13, fontWeight: 600, color: COLORS.textDark, marginBottom: 6 };

  return (
    <div>
      <h2 style={{ margin: "0 0 20px", fontSize: 20, fontWeight: 700, color: COLORS.textDark }}>
        🩸 血糖录入
      </h2>
      <SectionCard title="新增血糖记录">
        <form onSubmit={handleSubmit}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
            <div>
              <label style={labelStyle}>日期</label>
              <input
                type="date" style={inputStyle} value={form.date}
                onChange={(e) => setForm({ ...form, date: e.target.value })}
              />
            </div>
            <div>
              <label style={labelStyle}>时间</label>
              <input
                type="time" style={inputStyle} value={form.time}
                onChange={(e) => setForm({ ...form, time: e.target.value })}
              />
            </div>
            <div>
              <label style={labelStyle}>血糖值 (mmol/L)</label>
              <input
                type="number" step="0.1" min="1" max="35" placeholder="例如 6.5"
                style={{ ...inputStyle, fontSize: 18, fontWeight: 600 }}
                value={form.value}
                onChange={(e) => setForm({ ...form, value: e.target.value })}
              />
            </div>
            <div>
              <label style={labelStyle}>测量时段</label>
              <select
                style={inputStyle} value={form.mealTime}
                onChange={(e) => setForm({ ...form, mealTime: e.target.value })}
              >
                {MEAL_TIMES.map((m) => (
                  <option key={m.key} value={m.key}>{m.label}</option>
                ))}
              </select>
            </div>
          </div>
          <div style={{ marginBottom: 20 }}>
            <label style={labelStyle}>备注</label>
            <input
              type="text" placeholder="可选，例如：偷吃饼干、感觉头晕..."
              style={inputStyle} value={form.note}
              onChange={(e) => setForm({ ...form, note: e.target.value })}
            />
          </div>
          <button type="submit" style={{
            background: COLORS.glucoseLine, color: "#fff", border: "none",
            borderRadius: 8, padding: "12px 32px", fontSize: 15, fontWeight: 600,
            cursor: "pointer", transition: "opacity 0.2s",
          }}>
            保存记录
          </button>
        </form>
      </SectionCard>

      {/* Quick Reference */}
      <SectionCard title="📋 控制目标参考">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
          <div style={{ background: "#f0fdf4", borderRadius: 8, padding: 16 }}>
            <div style={{ fontWeight: 600, color: COLORS.success, marginBottom: 8 }}>空腹 / 餐前</div>
            <div style={{ fontSize: 24, fontWeight: 700, color: COLORS.success }}>&lt; 7.0 <span style={{ fontSize: 13, fontWeight: 400 }}>mmol/L</span></div>
          </div>
          <div style={{ background: "#fffbeb", borderRadius: 8, padding: 16 }}>
            <div style={{ fontWeight: 600, color: COLORS.warning, marginBottom: 8 }}>餐后 2 小时</div>
            <div style={{ fontSize: 24, fontWeight: 700, color: COLORS.warning }}>&lt; 10.0 <span style={{ fontSize: 13, fontWeight: 400 }}>mmol/L</span></div>
          </div>
        </div>
      </SectionCard>
    </div>
  );
}

// ============================================================
// BP ENTRY VIEW
// ============================================================

function BPEntryView({ onSave, showToast }) {
  const [form, setForm] = useState({
    date: new Date().toISOString().split("T")[0],
    time: new Date().toTimeString().slice(0, 5),
    systolic: "",
    diastolic: "",
    pulse: "",
    note: "",
  });

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!form.systolic || !form.diastolic || !form.pulse) {
      showToast("请填写完整的血压和脉搏数据", "error");
      return;
    }
    onSave({
      id: genId(),
      date: form.date,
      time: form.time,
      systolic: parseInt(form.systolic),
      diastolic: parseInt(form.diastolic),
      pulse: parseInt(form.pulse),
      note: form.note,
    });
    showToast("血压记录已保存", "success");
    setForm({ ...form, systolic: "", diastolic: "", pulse: "", note: "" });
  };

  const inputStyle = {
    width: "100%", padding: "10px 14px", border: `1px solid ${COLORS.border}`,
    borderRadius: 8, fontSize: 14, outline: "none", boxSizing: "border-box",
  };
  const labelStyle = { display: "block", fontSize: 13, fontWeight: 600, color: COLORS.textDark, marginBottom: 6 };

  return (
    <div>
      <h2 style={{ margin: "0 0 20px", fontSize: 20, fontWeight: 700, color: COLORS.textDark }}>
        💓 血压录入
      </h2>
      <SectionCard title="新增血压记录">
        <form onSubmit={handleSubmit}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
            <div>
              <label style={labelStyle}>日期</label>
              <input type="date" style={inputStyle} value={form.date}
                onChange={(e) => setForm({ ...form, date: e.target.value })} />
            </div>
            <div>
              <label style={labelStyle}>时间</label>
              <input type="time" style={inputStyle} value={form.time}
                onChange={(e) => setForm({ ...form, time: e.target.value })} />
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginBottom: 16 }}>
            <div>
              <label style={labelStyle}>高压 (mmHg)</label>
              <input type="number" min="50" max="250" placeholder="例如 120"
                style={{ ...inputStyle, fontSize: 18, fontWeight: 600, color: COLORS.bpSystolic }}
                value={form.systolic}
                onChange={(e) => setForm({ ...form, systolic: e.target.value })} />
            </div>
            <div>
              <label style={labelStyle}>低压 (mmHg)</label>
              <input type="number" min="30" max="150" placeholder="例如 80"
                style={{ ...inputStyle, fontSize: 18, fontWeight: 600, color: COLORS.bpDiastolic }}
                value={form.diastolic}
                onChange={(e) => setForm({ ...form, diastolic: e.target.value })} />
            </div>
            <div>
              <label style={labelStyle}>脉搏 (bpm)</label>
              <input type="number" min="30" max="200" placeholder="例如 72"
                style={{ ...inputStyle, fontSize: 18, fontWeight: 600, color: COLORS.pulse }}
                value={form.pulse}
                onChange={(e) => setForm({ ...form, pulse: e.target.value })} />
            </div>
          </div>
          <div style={{ marginBottom: 20 }}>
            <label style={labelStyle}>备注</label>
            <input type="text" placeholder="可选" style={inputStyle} value={form.note}
              onChange={(e) => setForm({ ...form, note: e.target.value })} />
          </div>
          <button type="submit" style={{
            background: COLORS.bpSystolic, color: "#fff", border: "none",
            borderRadius: 8, padding: "12px 32px", fontSize: 15, fontWeight: 600,
            cursor: "pointer",
          }}>
            保存记录
          </button>
        </form>
      </SectionCard>

      <SectionCard title="📋 正常范围参考">
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
          <div style={{ background: "#fef2f2", borderRadius: 8, padding: 16 }}>
            <div style={{ fontWeight: 600, color: COLORS.bpSystolic, marginBottom: 8 }}>高压</div>
            <div style={{ fontSize: 20, fontWeight: 700 }}>90-140 <span style={{ fontSize: 12, fontWeight: 400 }}>mmHg</span></div>
          </div>
          <div style={{ background: "#eff6ff", borderRadius: 8, padding: 16 }}>
            <div style={{ fontWeight: 600, color: COLORS.bpDiastolic, marginBottom: 8 }}>低压</div>
            <div style={{ fontSize: 20, fontWeight: 700 }}>60-90 <span style={{ fontSize: 12, fontWeight: 400 }}>mmHg</span></div>
          </div>
          <div style={{ background: "#f5f3ff", borderRadius: 8, padding: 16 }}>
            <div style={{ fontWeight: 600, color: COLORS.pulse, marginBottom: 8 }}>脉搏</div>
            <div style={{ fontSize: 20, fontWeight: 700 }}>60-100 <span style={{ fontSize: 12, fontWeight: 400 }}>bpm</span></div>
          </div>
        </div>
      </SectionCard>
    </div>
  );
}

// ============================================================
// IMPORT VIEW
// ============================================================

function ImportView({ onImportGlucose, onImportBP, showToast }) {
  const [file, setFile] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState(null);
  const [importType, setImportType] = useState("glucose");
  const [preview, setPreview] = useState([]);
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);

  const handleFileSelect = (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setFile(f);
    setSheets([]);
    setPreview([]);

    try {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const data = evt.target.result;
        const wb = XLSX.read(data, { type: "array" });
      const sheetList = wb.SheetNames.map((name) => {
        const ws = wb.Sheets[name];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
        let type = "unknown";
        const headerStr = headers.join(",").toLowerCase();
        if (headerStr.includes("血糖") || headerStr.includes("value") || headerStr.includes("mmol")) {
          type = "glucose";
        } else if (headerStr.includes("高压") || headerStr.includes("systolic") || headerStr.includes("脉搏")) {
          type = "bp";
        }
        return { name, rows, headers, type, count: rows.length };
      });
      setSheets(sheetList);
      showToast(`文件解析成功，发现 ${sheetList.length} 个工作表`, "success");
      };
      reader.readAsArrayBuffer(f);
    } catch (err) {
      showToast("文件解析失败: " + err.message, "error");
    }
  };

  const handlePreview = (sheet) => {
    setSelectedSheet(sheet);
    setImportType(sheet.type === "bp" ? "bp" : "glucose");
    setPreview(sheet.rows.slice(0, 5));
  };

  const handleImport = () => {
    if (!selectedSheet) return;
    setImporting(true);
    const rows = selectedSheet.rows;
    let imported = 0;

    if (importType === "glucose") {
      const records = rows
        .filter((r) => r["日期"] || r["date"])
        .map((r) => ({
          id: genId(),
          date: (r["日期"] || r["date"] || "").toString().replace(/\//g, "-"),
          time: (r["时间"] || r["time"] || "00:00").toString().slice(0, 5),
          value: parseFloat(r["血糖值 mmol/L"] || r["血糖值"] || r["value"] || 0),
          mealTime: "random",
          note: (r["备注"] || r["note"] || "").toString(),
        }))
        .filter((r) => r.value > 0);
      records.forEach((r) => onImportGlucose(r));
      imported = records.length;
    } else {
      const records = rows
        .filter((r) => r["日期"] || r["date"])
        .map((r) => ({
          id: genId(),
          date: (r["日期"] || r["date"] || "").toString().replace(/\//g, "-"),
          time: (r["时间"] || r["time"] || "00:00").toString().slice(0, 5),
          systolic: parseInt(r["高压 mmHg"] || r["高压"] || r["systolic"] || 0),
          diastolic: parseInt(r["低压 mmHg"] || r["低压"] || r["diastolic"] || 0),
          pulse: parseInt(r["脉搏 bpm"] || r["脉搏"] || r["pulse"] || 0),
          note: (r["备注"] || r["note"] || "").toString(),
        }))
        .filter((r) => r.systolic > 0);
      records.forEach((r) => onImportBP(r));
      imported = records.length;
    }

    setTimeout(() => {
      setImporting(false);
      showToast(`成功导入 ${imported} 条${importType === "glucose" ? "血糖" : "血压"}记录`, "success");
      setFile(null);
      setSheets([]);
      setSelectedSheet(null);
      setPreview([]);
    }, 500);
  };

  const cardStyle = {
    background: COLORS.cardBg, borderRadius: 12, padding: 24,
    boxShadow: "0 1px 3px rgba(0,0,0,0.08)", border: `1px solid ${COLORS.border}`, marginBottom: 20,
  };

  return (
    <div>
      <h2 style={{ margin: "0 0 20px", fontSize: 20, fontWeight: 700, color: COLORS.textDark }}>
        📁 文件导入
      </h2>

      <div style={cardStyle}>
        <h3 style={{ margin: "0 0 16px", fontSize: 16, fontWeight: 600 }}>上传 Excel 文件</h3>
        <div
          onClick={() => fileInputRef.current?.click()}
          style={{
            border: `2px dashed ${COLORS.border}`, borderRadius: 12,
            padding: "48px 20px", textAlign: "center", cursor: "pointer",
            background: "#f8fafc", transition: "background 0.2s",
          }}
          onMouseEnter={(e) => (e.currentTarget.style.background = "#f1f5f9")}
          onMouseLeave={(e) => (e.currentTarget.style.background = "#f8fafc")}
        >
          <div style={{ fontSize: 40, marginBottom: 12 }}>📂</div>
          <div style={{ fontSize: 15, fontWeight: 500, color: COLORS.textDark }}>
            {file ? file.name : "点击选择 .xlsx 文件"}
          </div>
          <div style={{ fontSize: 12, color: COLORS.textMid, marginTop: 4 }}>
            支持 Excel 格式（.xlsx）
          </div>
          <input
            ref={fileInputRef} type="file" accept=".xlsx,.xls"
            style={{ display: "none" }} onChange={handleFileSelect}
          />
        </div>
      </div>

      {/* Sheet List */}
      {sheets.length > 0 && (
        <div style={cardStyle}>
          <h3 style={{ margin: "0 0 16px", fontSize: 16, fontWeight: 600 }}>选择工作表</h3>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {sheets.map((sheet, i) => (
              <div
                key={i}
                onClick={() => handlePreview(sheet)}
                style={{
                  padding: "14px 18px", borderRadius: 8,
                  border: `1px solid ${selectedSheet?.name === sheet.name ? COLORS.primary : COLORS.border}`,
                  background: selectedSheet?.name === sheet.name ? "#eff6ff" : "#fff",
                  cursor: "pointer", transition: "all 0.2s",
                }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <span style={{ fontWeight: 600, fontSize: 14 }}>{sheet.name}</span>
                    <span style={{
                      marginLeft: 10, fontSize: 11, padding: "2px 8px", borderRadius: 10,
                      background: sheet.type === "glucose" ? "#fef3c7" : sheet.type === "bp" ? "#dbeafe" : "#f1f5f9",
                      color: sheet.type === "glucose" ? COLORS.warning : sheet.type === "bp" ? COLORS.primary : COLORS.textMid,
                    }}>
                      {sheet.type === "glucose" ? "血糖数据" : sheet.type === "bp" ? "血压数据" : "未知类型"}
                    </span>
                  </div>
                  <span style={{ fontSize: 12, color: COLORS.textMid }}>{sheet.count} 行</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Preview & Import */}
      {selectedSheet && preview.length > 0 && (
        <div style={cardStyle}>
          <h3 style={{ margin: "0 0 12px", fontSize: 16, fontWeight: 600 }}>数据预览</h3>
          <div style={{ marginBottom: 16 }}>
            <label style={{ fontSize: 13, fontWeight: 600, marginRight: 12 }}>导入为：</label>
            <select
              value={importType}
              onChange={(e) => setImportType(e.target.value)}
              style={{ padding: "6px 12px", borderRadius: 6, border: `1px solid ${COLORS.border}`, fontSize: 13 }}
            >
              <option value="glucose">血糖数据</option>
              <option value="bp">血压数据</option>
            </select>
          </div>
          <div style={{ overflowX: "auto", marginBottom: 16 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr>
                  {selectedSheet.headers.map((h, i) => (
                    <th key={i} style={{
                      padding: "8px 12px", textAlign: "left", borderBottom: `2px solid ${COLORS.border}`,
                      background: "#f8fafc", fontWeight: 600, whiteSpace: "nowrap",
                    }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.map((row, i) => (
                  <tr key={i}>
                    {selectedSheet.headers.map((h, j) => (
                      <td key={j} style={{
                        padding: "8px 12px", borderBottom: `1px solid ${COLORS.border}`,
                        whiteSpace: "nowrap",
                      }}>{String(row[h] || "")}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <button
            onClick={handleImport}
            disabled={importing}
            style={{
              background: importing ? COLORS.textLight : COLORS.primary,
              color: "#fff", border: "none", borderRadius: 8,
              padding: "12px 32px", fontSize: 15, fontWeight: 600,
              cursor: importing ? "not-allowed" : "pointer",
            }}
          >
            {importing ? "导入中..." : `确认导入 (${selectedSheet.count} 条)`}
          </button>
        </div>
      )}
    </div>
  );
}

// ============================================================
// HISTORY VIEW
// ============================================================

function HistoryView({ glucose, bp, onDeleteGlucose, onDeleteBP }) {
  const [viewType, setViewType] = useState("glucose");
  const [sortDir, setSortDir] = useState("desc");

  const sortedGlucose = useMemo(() => {
    return [...glucose].sort((a, b) => {
      const cmp = (a.date + a.time).localeCompare(b.date + b.time);
      return sortDir === "desc" ? -cmp : cmp;
    });
  }, [glucose, sortDir]);

  const sortedBP = useMemo(() => {
    return [...bp].sort((a, b) => {
      const cmp = (a.date + a.time).localeCompare(b.date + b.time);
      return sortDir === "desc" ? -cmp : cmp;
    });
  }, [bp, sortDir]);

  const thStyle = {
    padding: "10px 14px", textAlign: "left", borderBottom: `2px solid ${COLORS.border}`,
    background: "#f8fafc", fontWeight: 600, fontSize: 13, whiteSpace: "nowrap",
  };
  const tdStyle = {
    padding: "10px 14px", borderBottom: `1px solid ${COLORS.border}`, fontSize: 13,
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: COLORS.textDark }}>📋 历史记录</h2>
        <div style={{ display: "flex", gap: 8 }}>
          <button
            onClick={() => setViewType("glucose")}
            style={{
              padding: "8px 16px", borderRadius: 8, border: `1px solid ${COLORS.border}`,
              background: viewType === "glucose" ? COLORS.glucoseLine : "#fff",
              color: viewType === "glucose" ? "#fff" : COLORS.textDark,
              fontWeight: 600, fontSize: 13, cursor: "pointer",
            }}
          >
            血糖 ({glucose.length})
          </button>
          <button
            onClick={() => setViewType("bp")}
            style={{
              padding: "8px 16px", borderRadius: 8, border: `1px solid ${COLORS.border}`,
              background: viewType === "bp" ? COLORS.bpSystolic : "#fff",
              color: viewType === "bp" ? "#fff" : COLORS.textDark,
              fontWeight: 600, fontSize: 13, cursor: "pointer",
            }}
          >
            血压 ({bp.length})
          </button>
          <button
            onClick={() => setSortDir(sortDir === "desc" ? "asc" : "desc")}
            style={{
              padding: "8px 16px", borderRadius: 8, border: `1px solid ${COLORS.border}`,
              background: "#fff", fontSize: 13, cursor: "pointer",
            }}
          >
            {sortDir === "desc" ? "最新在前 ↓" : "最早在前 ↑"}
          </button>
        </div>
      </div>

      <div style={{
        background: COLORS.cardBg, borderRadius: 12, overflow: "hidden",
        boxShadow: "0 1px 3px rgba(0,0,0,0.08)", border: `1px solid ${COLORS.border}`,
      }}>
        {viewType === "glucose" ? (
          sortedGlucose.length === 0 ? (
            <EmptyState icon="🩸" title="暂无血糖记录" desc="去录入页面添加第一条记录" />
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th style={thStyle}>日期</th>
                    <th style={thStyle}>时间</th>
                    <th style={thStyle}>血糖值</th>
                    <th style={thStyle}>时段</th>
                    <th style={thStyle}>状态</th>
                    <th style={thStyle}>备注</th>
                    <th style={thStyle}>操作</th>
                  </tr>
                </thead>
                <tbody>
                  {sortedGlucose.map((r) => {
                    const status = classifyGlucose(r.value, r.mealTime);
                    return (
                      <tr key={r.id} style={{ transition: "background 0.15s" }}
                        onMouseEnter={(e) => (e.currentTarget.style.background = "#f8fafc")}
                        onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                      >
                        <td style={tdStyle}>{r.date}</td>
                        <td style={tdStyle}>{r.time}</td>
                        <td style={{ ...tdStyle, fontWeight: 700, fontSize: 15, color: COLORS[status] }}>{r.value}</td>
                        <td style={tdStyle}>{getMealLabel(r.mealTime)}</td>
                        <td style={tdStyle}>
                          <span style={{
                            padding: "3px 10px", borderRadius: 12, fontSize: 12, fontWeight: 600,
                            background: status === "达标" ? "#f0fdf4" : status === "超标" ? "#fef2f2" : "#fffbeb",
                            color: COLORS[status],
                          }}>{status}</span>
                        </td>
                        <td style={{ ...tdStyle, color: COLORS.textMid, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis" }}>
                          {r.note || "-"}
                        </td>
                        <td style={tdStyle}>
                          <button
                            onClick={() => onDeleteGlucose(r.id)}
                            style={{
                              background: "none", border: "none", color: COLORS.danger,
                              cursor: "pointer", fontSize: 13, padding: "4px 8px",
                            }}
                          >删除</button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )
        ) : (
          sortedBP.length === 0 ? (
            <EmptyState icon="💓" title="暂无血压记录" desc="去录入页面添加第一条记录" />
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th style={thStyle}>日期</th>
                    <th style={thStyle}>时间</th>
                    <th style={thStyle}>高压</th>
                    <th style={thStyle}>低压</th>
                    <th style={thStyle}>脉搏</th>
                    <th style={thStyle}>状态</th>
                    <th style={thStyle}>备注</th>
                    <th style={thStyle}>操作</th>
                  </tr>
                </thead>
                <tbody>
                  {sortedBP.map((r) => {
                    const status = classifyBP(r.systolic, r.diastolic);
                    return (
                      <tr key={r.id}
                        onMouseEnter={(e) => (e.currentTarget.style.background = "#f8fafc")}
                        onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                      >
                        <td style={tdStyle}>{r.date}</td>
                        <td style={tdStyle}>{r.time}</td>
                        <td style={{ ...tdStyle, fontWeight: 700, color: COLORS.bpSystolic }}>{r.systolic}</td>
                        <td style={{ ...tdStyle, fontWeight: 700, color: COLORS.bpDiastolic }}>{r.diastolic}</td>
                        <td style={{ ...tdStyle, fontWeight: 600, color: COLORS.pulse }}>{r.pulse}</td>
                        <td style={tdStyle}>
                          <span style={{
                            padding: "3px 10px", borderRadius: 12, fontSize: 12, fontWeight: 600,
                            background: status === "达标" ? "#f0fdf4" : status === "偏高" ? "#fef2f2" : "#eff6ff",
                            color: COLORS[status] || COLORS.primary,
                          }}>{status}</span>
                        </td>
                        <td style={{ ...tdStyle, color: COLORS.textMid }}>{r.note || "-"}</td>
                        <td style={tdStyle}>
                          <button
                            onClick={() => onDeleteBP(r.id)}
                            style={{
                              background: "none", border: "none", color: COLORS.danger,
                              cursor: "pointer", fontSize: 13, padding: "4px 8px",
                            }}
                          >删除</button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )
        )}
      </div>
    </div>
  );
}

// ============================================================
// ANALYSIS VIEW
// ============================================================

function AnalysisView({ glucose, bp }) {
  // Before/after medication stop comparison
  const medStopDate = "2026-06-10";

  const beforeGlucose = glucose.filter((r) => r.date < medStopDate);
  const afterGlucose = glucose.filter((r) => r.date >= medStopDate);
  const beforeBP = bp.filter((r) => r.date < medStopDate);
  const afterBP = bp.filter((r) => r.date >= medStopDate);

  const beforeGlucoseStats = calcStats(beforeGlucose, "value");
  const afterGlucoseStats = calcStats(afterGlucose, "value");
  const beforeBPStats = calcStats(beforeBP, "systolic");
  const afterBPStats = calcStats(afterBP, "systolic");

  const before达标率 = calc达标率(beforeGlucose, (r) => classifyGlucose(r.value, r.mealTime));
  const after达标率 = calc达标率(afterGlucose, (r) => classifyGlucose(r.value, r.mealTime));

  // Consecutive 达标 days
  const consecutive达标 = useMemo(() => {
    const byDate = {};
    glucose.forEach((r) => {
      if (!byDate[r.date]) byDate[r.date] = [];
      byDate[r.date].push(r);
    });
    const dates = Object.keys(byDate).sort();
    let max = 0, current = 0;
    dates.forEach((d) => {
      const all达标 = byDate[d].every((r) => classifyGlucose(r.value, r.mealTime) === "达标");
      if (all达标) { current++; max = Math.max(max, current); }
      else { current = 0; }
    });
    return max;
  }, [glucose]);

  // Problem time slots
  const problemSlots = useMemo(() => {
    const groups = {};
    MEAL_TIMES.forEach((m) => { groups[m.key] = { total: 0, 超标: 0, values: [] }; });
    glucose.forEach((r) => {
      const key = r.mealTime || "random";
      if (!groups[key]) return;
      groups[key].total++;
      groups[key].values.push(r.value);
      if (classifyGlucose(r.value, r.mealTime) === "超标") groups[key].超标++;
    });
    return MEAL_TIMES
      .map((m) => ({
        name: m.label,
        key: m.key,
        total: groups[m.key]?.total || 0,
        超标: groups[m.key]?.超标 || 0,
        超标率: groups[m.key]?.total > 0
          ? Math.round((groups[m.key].超标 / groups[m.key].total) * 100)
          : 0,
        avg: groups[m.key]?.values.length > 0
          ? +(groups[m.key].values.reduce((a, b) => a + b, 0) / groups[m.key].values.length).toFixed(1)
          : 0,
      }))
      .filter((s) => s.total > 0)
      .sort((a, b) => b.超标率 - a.超标率);
  }, [glucose]);

  // Weekly trend
  const weeklyTrend = useMemo(() => {
    const byWeek = {};
    glucose.forEach((r) => {
      const d = new Date(r.date);
      const weekStart = new Date(d);
      weekStart.setDate(d.getDate() - d.getDay());
      const key = weekStart.toISOString().split("T")[0];
      if (!byWeek[key]) byWeek[key] = [];
      byWeek[key].push(r.value);
    });
    return Object.entries(byWeek)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([week, values]) => ({
        week: `第${Object.keys(byWeek).sort().indexOf(week) + 1}周`,
        avg: +(values.reduce((a, b) => a + b, 0) / values.length).toFixed(1),
        count: values.length,
      }));
  }, [glucose]);

  const cardStyle = {
    background: COLORS.cardBg, borderRadius: 12, padding: 24,
    boxShadow: "0 1px 3px rgba(0,0,0,0.08)", border: `1px solid ${COLORS.border}`, marginBottom: 20,
  };

  return (
    <div>
      <h2 style={{ margin: "0 0 20px", fontSize: 20, fontWeight: 700, color: COLORS.textDark }}>
        🔍 深度分析
      </h2>

      {/* Key Metrics */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 16, marginBottom: 24 }}>
        <StatCard title="连续达标天数" value={consecutive达标} unit="天" icon="🔥" color={COLORS.success} />
        <StatCard title="最低血糖" value={calcStats(glucose, "value").min} unit="mmol/L" icon="⬇️" color={COLORS.success} />
        <StatCard title="最高血糖" value={calcStats(glucose, "value").max} unit="mmol/L" icon="⬆️" color={COLORS.danger} />
        <StatCard title="总记录天数" value={new Set(glucose.map(r => r.date)).size} unit="天" icon="📅" color={COLORS.primary} />
      </div>

      {/* Before/After Medication Stop */}
      <div style={cardStyle}>
        <h3 style={{ margin: "0 0 16px", fontSize: 16, fontWeight: 600 }}>
          💊 停药前后对比（6月10日停用伊伐布雷定+美托洛尔）
        </h3>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
          <div>
            <h4 style={{ margin: "0 0 12px", fontSize: 14, color: COLORS.textMid }}>血糖对比</h4>
            <div style={{ display: "flex", gap: 16 }}>
              <div style={{ flex: 1, background: "#f8fafc", borderRadius: 8, padding: 14, textAlign: "center" }}>
                <div style={{ fontSize: 12, color: COLORS.textMid, marginBottom: 4 }}>停药前</div>
                <div style={{ fontSize: 24, fontWeight: 700, color: COLORS.glucoseLine }}>{beforeGlucoseStats.avg}</div>
                <div style={{ fontSize: 11, color: COLORS.textMid }}>达标率 {before达标率}%</div>
              </div>
              <div style={{ flex: 1, background: "#f8fafc", borderRadius: 8, padding: 14, textAlign: "center" }}>
                <div style={{ fontSize: 12, color: COLORS.textMid, marginBottom: 4 }}>停药后</div>
                <div style={{ fontSize: 24, fontWeight: 700, color: COLORS.glucoseLine }}>{afterGlucoseStats.avg}</div>
                <div style={{ fontSize: 11, color: COLORS.textMid }}>达标率 {after达标率}%</div>
              </div>
            </div>
          </div>
          <div>
            <h4 style={{ margin: "0 0 12px", fontSize: 14, color: COLORS.textMid }}>血压（高压）对比</h4>
            <div style={{ display: "flex", gap: 16 }}>
              <div style={{ flex: 1, background: "#f8fafc", borderRadius: 8, padding: 14, textAlign: "center" }}>
                <div style={{ fontSize: 12, color: COLORS.textMid, marginBottom: 4 }}>停药前</div>
                <div style={{ fontSize: 24, fontWeight: 700, color: COLORS.bpSystolic }}>{beforeBPStats.avg}</div>
                <div style={{ fontSize: 11, color: COLORS.textMid }}>mmHg</div>
              </div>
              <div style={{ flex: 1, background: "#f8fafc", borderRadius: 8, padding: 14, textAlign: "center" }}>
                <div style={{ fontSize: 12, color: COLORS.textMid, marginBottom: 4 }}>停药后</div>
                <div style={{ fontSize: 24, fontWeight: 700, color: COLORS.bpSystolic }}>{afterBPStats.avg}</div>
                <div style={{ fontSize: 11, color: COLORS.textMid }}>mmHg</div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Problem Time Slots */}
      <div style={cardStyle}>
        <h3 style={{ margin: "0 0 16px", fontSize: 16, fontWeight: 600 }}>⚠️ 血糖控制薄弱时段</h3>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {problemSlots.map((slot) => (
            <div key={slot.key} style={{
              display: "flex", alignItems: "center", gap: 16,
              padding: "12px 16px", borderRadius: 8,
              background: slot.超标率 > 50 ? "#fef2f2" : slot.超标率 > 20 ? "#fffbeb" : "#f0fdf4",
            }}>
              <div style={{ fontWeight: 600, fontSize: 14, minWidth: 80 }}>{slot.name}</div>
              <div style={{ flex: 1, height: 8, background: "#e2e8f0", borderRadius: 4, overflow: "hidden" }}>
                <div style={{
                  height: "100%", borderRadius: 4,
                  width: `${slot.超标率}%`,
                  background: slot.超标率 > 50 ? COLORS.danger : slot.超标率 > 20 ? COLORS.warning : COLORS.success,
                }} />
              </div>
              <div style={{ fontSize: 13, fontWeight: 600, minWidth: 60, textAlign: "right",
                color: slot.超标率 > 50 ? COLORS.danger : slot.超标率 > 20 ? COLORS.warning : COLORS.success
              }}>
                超标 {slot.超标率}%
              </div>
              <div style={{ fontSize: 12, color: COLORS.textMid, minWidth: 80 }}>
                均值 {slot.avg} ({slot.total}次)
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Weekly Trend */}
      <SectionCard title="📊 周均值趋势">
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={weeklyTrend}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
            <XAxis dataKey="week" tick={{ fontSize: 12, fill: COLORS.textMid }} />
            <YAxis domain={[4, 10]} tick={{ fontSize: 12, fill: COLORS.textMid }} />
            <Tooltip contentStyle={{ fontSize: 13, borderRadius: 8 }} />
            <ReferenceLine y={TARGETS.glucose.fastingMax} stroke={COLORS.success} strokeDasharray="5 5" />
            <Bar dataKey="avg" name="周均值" fill={COLORS.glucoseLine} radius={[6, 6, 0, 0]}>
              {weeklyTrend.map((entry, index) => (
                <Cell
                  key={index}
                  fill={entry.avg <= TARGETS.glucose.fastingMax ? COLORS.success :
                    entry.avg <= TARGETS.glucose.postMealMax ? COLORS.warning : COLORS.danger}
                />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </SectionCard>
    </div>
  );
}

// ============================================================
// MAIN COMPONENT
// ============================================================

export default function HealthMonitorDashboard() {
  const [activeTab, setActiveTab] = useState("dashboard");
  const [glucose, setGlucose] = useState(SAMPLE_GLUCOSE);
  const [bp, setBP] = useState(SAMPLE_BP);
  const [toast, setToast] = useState(null);

  const showToast = useCallback((message, type = "info") => {
    setToast({ message, type });
  }, []);

  const addGlucose = useCallback((record) => {
    setGlucose((prev) => [...prev, record]);
  }, []);

  const addBP = useCallback((record) => {
    setBP((prev) => [...prev, record]);
  }, []);

  const deleteGlucose = useCallback((id) => {
    setGlucose((prev) => prev.filter((r) => r.id !== id));
    showToast("记录已删除", "success");
  }, [showToast]);

  const deleteBP = useCallback((id) => {
    setBP((prev) => prev.filter((r) => r.id !== id));
    showToast("记录已删除", "success");
  }, [showToast]);

  const renderView = () => {
    switch (activeTab) {
      case "dashboard":
        return <DashboardView glucose={glucose} bp={bp} />;
      case "glucose":
        return <GlucoseEntryView onSave={addGlucose} showToast={showToast} />;
      case "bp":
        return <BPEntryView onSave={addBP} showToast={showToast} />;
      case "import":
        return (
          <ImportView
            onImportGlucose={addGlucose}
            onImportBP={addBP}
            showToast={showToast}
          />
        );
      case "history":
        return (
          <HistoryView
            glucose={glucose} bp={bp}
            onDeleteGlucose={deleteGlucose}
            onDeleteBP={deleteBP}
          />
        );
      case "analysis":
        return <AnalysisView glucose={glucose} bp={bp} />;
      default:
        return <DashboardView glucose={glucose} bp={bp} />;
    }
  };

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif", background: COLORS.pageBg }}>
      {/* Sidebar */}
      <div style={{
        width: 210, minWidth: 210, background: COLORS.sidebarBg,
        color: "#fff", display: "flex", flexDirection: "column",
        padding: "0", overflow: "hidden",
      }}>
        {/* Logo area */}
        <div style={{
          padding: "24px 20px 20px",
          borderBottom: "1px solid #1e293b",
        }}>
          <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 4 }}>
            ❤️ 健康监测
          </div>
          <div style={{ fontSize: 12, color: "#94a3b8" }}>
            血糖 · 血压 · 脉搏
          </div>
        </div>

        {/* Nav items */}
        <div style={{ flex: 1, padding: "12px 8px", overflowY: "auto" }}>
          {TABS.map((tab) => (
            <div
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              style={{
                padding: "11px 14px", borderRadius: 8, marginBottom: 4,
                cursor: "pointer", display: "flex", alignItems: "center", gap: 10,
                background: activeTab === tab.id ? COLORS.sidebarActive : "transparent",
                color: activeTab === tab.id ? "#fff" : "#94a3b8",
                fontWeight: activeTab === tab.id ? 600 : 400,
                fontSize: 14, transition: "all 0.15s",
              }}
              onMouseEnter={(e) => {
                if (activeTab !== tab.id) e.currentTarget.style.background = COLORS.sidebarHover;
              }}
              onMouseLeave={(e) => {
                if (activeTab !== tab.id) e.currentTarget.style.background = "transparent";
              }}
            >
              <span>{tab.icon}</span>
              <span>{tab.label}</span>
            </div>
          ))}
        </div>

        {/* Footer */}
        <div style={{
          padding: "16px 20px", borderTop: "1px solid #1e293b",
          fontSize: 11, color: "#64748b",
        }}>
          <div>2型糖尿病 · 达格列净+二甲双胍</div>
          <div style={{ marginTop: 4 }}>最近更新: 2026-06-18</div>
        </div>
      </div>

      {/* Main content */}
      <div style={{ flex: 1, overflow: "auto", padding: 28 }}>
        {renderView()}
      </div>

      {/* Toast */}
      {toast && (
        <Toast
          message={toast.message}
          type={toast.type}
          onClose={() => setToast(null)}
        />
      )}
    </div>
  );
}
