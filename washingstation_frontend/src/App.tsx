import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
{showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
import {
  BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from 'recharts'

// ─── Types ────────────────────────────────────────────────────────────────────

type Page =
  | 'login' | 'dashboard' | 'earnings' | 'pending-payments'
  | 'workers' | 'worker-detail' | 'add-transaction'
  | 'salary-settlement' | 'expenses' | 'reports' | 'settings'

interface Worker {
  id: number
  name: string
  salary: number
  joinDate: string
  initials: string
  color: string
  status: "Active" | "Inactive"
}

interface WTransaction {
  id: number
  workerId: number
  workerName: string
  amount: number
  date: string
  type: string
  remarks: string
}

interface Earning {
  id: number
  date: string
  customerName: string
  vehicleType: string
  washType: string
  amount: number
  paymentMode: 'Cash' | 'UPI' | 'Pending'
  status: 'Active' | 'Paid'
  paidDate?: string
  paidMode?: 'Cash' | 'UPI'
  remarks?: string
}

interface SettlementRecord {
  workerId: number
  month: number
  year: number
  settled: boolean
  settlementDate?: string
}

interface Expense {
  id: number;
  date: string;
  category: string;
  description: string;
  amount: number;
  paymentMode: string;
}

// ─── Initial Data ─────────────────────────────────────────────────────────────

const initialWorkers: Worker[] = [
  { id: 1, name: 'Ravi Kumar',   salary: 12000, joinDate: '2024-01-15', initials: 'RK', color: '#2563EB',status:'Active' },
  { id: 2, name: 'Mahesh Singh', salary: 11000, joinDate: '2024-03-01', initials: 'MS', color: '#F97316',status:'Active' },
  { id: 3, name: 'Suresh Patel', salary: 13000, joinDate: '2023-11-20', initials: 'SP', color: '#10B981',status:'Active' },
  { id: 4, name: 'Dinesh Yadav', salary: 10000, joinDate: '2024-06-10', initials: 'DY', color: '#8B5CF6',status:'Active'},
]

const initialEarnings: Earning[] = [
  { id: 1, date: '2025-07-19', customerName: '',        vehicleType: 'Car',  washType: 'Full Wash',    amount: 450, paymentMode: 'UPI',     status: 'Active' },
  { id: 2, date: '2025-07-19', customerName: '',        vehicleType: 'Bike', washType: 'Basic Wash',   amount: 150, paymentMode: 'Cash',    status: 'Active' },
  { id: 3, date: '2025-07-19', customerName: '',        vehicleType: 'Car',  washType: 'Premium Wash', amount: 650, paymentMode: 'UPI',     status: 'Active' },
  { id: 4, date: '2025-07-19', customerName: '',        vehicleType: 'SUV',  washType: 'Full Wash',    amount: 550, paymentMode: 'Cash',    status: 'Active' },
  { id: 5, date: '2025-07-18', customerName: '',        vehicleType: 'Car',  washType: 'Full Wash',    amount: 450, paymentMode: 'Cash',    status: 'Active' },
  { id: 6, date: '2025-07-18', customerName: '',        vehicleType: 'Car',  washType: 'Basic Wash',   amount: 250, paymentMode: 'UPI',     status: 'Active' },
  { id: 7, date: '2025-07-19', customerName: 'Kiran Patel',  vehicleType: 'Car',  washType: 'Full Wash',    amount: 450, paymentMode: 'Pending', status: 'Active', remarks: 'Will pay tomorrow' },
  { id: 8, date: '2025-07-18', customerName: 'Anita Sharma', vehicleType: 'SUV',  washType: 'Premium Wash', amount: 750, paymentMode: 'Pending', status: 'Active', remarks: '' },
  { id: 9, date: '2025-07-17', customerName: 'Vijay Modi',   vehicleType: 'Car',  washType: 'Full Wash',    amount: 450, paymentMode: 'Pending', status: 'Active', remarks: 'Regular customer' },
]

const initialSettlements: SettlementRecord[] = [
  { workerId: 1, month: 6, year: 2025, settled: true, settlementDate: '2025-06-30' },
  { workerId: 2, month: 6, year: 2025, settled: true, settlementDate: '2025-06-28' },
  { workerId: 3, month: 6, year: 2025, settled: true, settlementDate: '2025-06-29' },
  { workerId: 4, month: 6, year: 2025, settled: true, settlementDate: '2025-06-30' },
]


// ─── Helpers ──────────────────────────────────────────────────────────────────

const fmt = (n?: number) => `₹${(n ?? 0).toLocaleString("en-IN")}`;
const now = new Date();
const todayStr = now.toISOString().split("T")[0];
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December']

function getDateStr() {
  return now.toLocaleDateString('en-IN', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })
}

function calcRemaining(worker: Worker, transactions: WTransaction[]) {
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  const total = transactions
    .filter(t => {
      if (t.workerId !== worker.id) return false;

      const d = new Date(t.date);

      return (
        d.getMonth() === currentMonth &&
        d.getFullYear() === currentYear
      );
    })
    .reduce((sum, t) => sum + t.amount, 0);

  return Math.max(0, worker.salary - total);
}

function calcDeductions(worker: Worker, transactions: WTransaction[]) {
  return transactions
    .filter(
      t =>
        t.workerId === worker.id &&
        ["Advance", "Food", "Rent", "Medical", "Bills"].includes(t.type)
    )
    .reduce((s, t) => s + t.amount, 0);
}

function txBadgeColor(type: string) {
  const map: Record<string, string> = {
    'Salary Settlement': 'bg-green-100 text-green-700',
    'Advance':     'bg-blue-100 text-blue-700',
    'Rent':        'bg-purple-100 text-purple-700',
    'Medical':     'bg-red-100 text-red-700',
    'Food':        'bg-yellow-100 text-yellow-700',
    'Bills':       'bg-orange-100 text-orange-700',
  }
  return map[type] ?? 'bg-gray-100 text-gray-700'
}

function activePending(earnings: Earning[]) {
  return earnings.filter(e => e.paymentMode === 'Pending' && e.status === 'Active')
}

// ─── Sidebar ──────────────────────────────────────────────────────────────────

const navItems = [
  { key: 'dashboard',         label: 'Dashboard',             emoji: '🏠' },
  { key: 'earnings',          label: 'Daily Earnings',        emoji: '💰' },
  { key: 'pending-payments',  label: 'Pending Payments',      emoji: '🧾' },
  { key: 'workers',           label: 'Workers',               emoji: '👷' },
  { key: 'add-transaction',   label: 'Worker Transactions',   emoji: '💸' },
  { key: 'salary-settlement', label: 'Salary Settlement',     emoji: '💳' },
  { key: 'expenses',          label: 'Expenses',              emoji: '💵' },
  { key: 'reports',           label: 'Reports',               emoji: '📊' },
  { key: 'settings',          label: 'Settings',              emoji: '⚙️' },
]

interface SidebarProps {
  current: Page
  onNav: (p: Page) => void
  collapsed: boolean
  onToggle: () => void
  mobileOpen: boolean
  onMobileClose: () => void
  onLogout: () => void
}

function Sidebar({ current, onNav, collapsed, onToggle, mobileOpen, onMobileClose, onLogout }: SidebarProps) {
  const nav = (key: string) => { onNav(key as Page); onMobileClose() }

  const inner = (
    <div className="flex flex-col h-full bg-white border-r border-gray-100 shadow-sm">
      <div className="flex items-center justify-between px-4 py-5 border-b border-gray-100 min-h-[72px]">
        {!collapsed && (
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white text-sm font-bold flex-shrink-0">W</div>
            <div className="leading-tight">
              <p className="font-bold text-gray-900 text-sm">Wash Station</p>
              <p className="text-blue-600 font-semibold text-xs">Management</p>
            </div>
          </div>
        )}
        <button onClick={onToggle} className="p-1.5 rounded-lg hover:bg-gray-100 transition-colors text-gray-500 ml-auto">
          <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
            <path strokeLinecap="round" d="M4 6h16M4 12h16M4 18h16" />
          </svg>
        </button>
      </div>

      <nav className="flex-1 py-4 px-2 space-y-0.5 overflow-y-auto">
        {navItems.map(item => {
          const active = current === item.key
          return (
            <button key={item.key} onClick={() => nav(item.key)} title={collapsed ? item.label : undefined}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium transition-all ${
                active ? 'bg-blue-600 text-white shadow-sm' : 'text-gray-600 hover:bg-blue-50 hover:text-blue-700'
              } ${collapsed ? 'justify-center' : ''}`}>
              <span className="text-base flex-shrink-0">{item.emoji}</span>
              {!collapsed && <span className="truncate">{item.label}</span>}
            </button>
          )
        })}
      </nav>

      <div className="p-2 border-t border-gray-100">
        <button onClick={onLogout}
          className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium text-red-500 hover:bg-red-50 transition-all ${collapsed ? 'justify-center' : ''}`}>
          <span className="text-base">🚪</span>
          {!collapsed && <span>Logout</span>}
        </button>
      </div>
    </div>
  )

  return (
    <>
      <aside className={`hidden lg:flex flex-col h-screen sticky top-0 transition-all duration-300 flex-shrink-0 ${collapsed ? 'w-[72px]' : 'w-64'}`}>
        {inner}
      </aside>
      {mobileOpen && (
        <>
          <div className="fixed inset-0 bg-black/40 z-40 lg:hidden" onClick={onMobileClose} />
          <aside className="fixed left-0 top-0 h-full w-64 z-50 lg:hidden">{inner}</aside>
        </>
      )}
    </>
  )
}

// ─── Top Bar ──────────────────────────────────────────────────────────────────

function TopBar({ title, onMenuClick }: { title: string; onMenuClick: () => void }) {
  return (
    <header className="bg-white border-b border-gray-100 px-4 py-3 flex items-center justify-between lg:px-6 flex-shrink-0">
      <div className="flex items-center gap-3">
        <button onClick={onMenuClick} className="lg:hidden p-2 rounded-lg hover:bg-gray-100 text-gray-600 transition-colors">
          <svg width="20" height="20" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
            <path strokeLinecap="round" d="M4 6h16M4 12h16M4 18h16" />
          </svg>
        </button>
        <h1 className="font-semibold text-gray-900 text-base lg:text-lg">{title}</h1>
      </div>
    </header>
  )
}

// ─── Stat Card ────────────────────────────────────────────────────────────────

function StatCard({ label, value, sub, emoji, bgColor }: { label: string; value: string; sub: string; emoji: string; bgColor: string }) {
  return (
    <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4 flex items-start justify-between gap-3">
      <div className="min-w-0">
        <p className="text-xs font-medium text-gray-500 mb-1 truncate">{label}</p>
        <p className="text-xl font-bold text-gray-900 truncate">{value}</p>
        <p className="text-xs text-gray-400 mt-1 truncate">{sub}</p>
      </div>
      <div className={`w-10 h-10 rounded-xl flex items-center justify-center text-lg flex-shrink-0 ${bgColor}`}>{emoji}</div>
    </div>
  )
}

// ─── Dashboard ────────────────────────────────────────────────────────────────

interface DashboardProps {
  earnings: Earning[]
  workers: Worker[]
  transactions: WTransaction[]
  expenses: Expense[]
  totalEarnings: number
  totalExpenses: number
  netProfit: number

  dashboard: {
    totalWorkers: number
    activeWorkers: number
    totalSalary: number
    totalTransactions: number
  }

  onNav: (p: Page) => void
  onSelectWorker: (w: Worker) => void
}

function DashboardPage({ earnings, workers, transactions, expenses, dashboard,totalEarnings,totalExpenses,netProfit, onNav, onSelectWorker }: DashboardProps) {
  const todayEntries = earnings.filter(e => e.date === todayStr && e.paymentMode !== 'Pending')
  const todayTotal   = todayEntries.reduce((s, e) => s + e.amount, 0)
  const todayCash    = todayEntries.filter(e => e.paymentMode === 'Cash').reduce((s, e) => s + e.amount, 0)
  const todayUPI     = todayEntries.filter(e => e.paymentMode === 'UPI').reduce((s, e) => s + e.amount, 0)
  const pending      = activePending(earnings)
  const pendingAmt   = pending.reduce((s, e) => s + e.amount, 0)
  const pendingSal   = workers.reduce((s, w) => s + calcRemaining(w, transactions), 0)
  console.table(transactions);
  console.log(
  transactions.filter(t => t.workerId === 3)
);

  const recentActivity = [
    ...transactions.slice(-4).reverse().map(t => ({
      label: `${t.workerName} – ${t.type}`,
      amount: fmt(t.amount),
      date: t.date,
      emoji: t.type === 'Salary Settlement' ? '✅' : t.type === 'Advance' ? '💙' : '📝',
      positive: false,
    })),
    ...earnings.filter(e => e.paymentMode !== 'Pending').slice(-4).reverse().map(e => ({
      label: `${e.vehicleType} Wash${e.customerName ? ` – ${e.customerName}` : ''}`,
      amount: `${fmt(e.amount)} · ${e.paymentMode}`,
      date: e.date,
      emoji: e.vehicleType === 'Bike' ? '🏍️' : e.vehicleType === 'SUV' ? '🚙' : '🚗',
      positive: true,
    })),
    ...pending.slice(-2).map(e => ({
      label: `⏳ Pending – ${e.customerName}`,
      amount: fmt(e.amount),
      date: e.date,
      emoji: '🧾',
      positive: false,
    })),
  ].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 7)

  return (
    <div className="p-4 lg:p-6 space-y-5">
      {/* Banner */}
      <div className="bg-gradient-to-r from-blue-600 to-blue-700 rounded-2xl p-5 text-white relative overflow-hidden">
        <div className="absolute -top-8 -right-8 w-40 h-40 rounded-full bg-white/5" />
        <div className="absolute bottom-0 right-16 w-20 h-20 rounded-full bg-orange-400/20 translate-y-4" />
        <p className="text-blue-200 text-sm mb-1 relative">👋 Welcome back, {localStorage.getItem("loggedUser")}</p>
        <h2 className="text-lg font-bold mb-1 relative">Wash Station Management System</h2>
        <p className="text-blue-200 text-sm relative">{getDateStr()}</p>
      </div>

      {/* Quick Actions */}
      <div>
        <h3 className="font-semibold text-gray-900 mb-3">Quick Actions</h3>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
          {[
            { label: '➕ Add Earning',         page: 'earnings'          as Page },
            { label: '🧾 Pending Payments',    page: 'pending-payments'  as Page },
            { label: '💸 Worker Transaction',  page: 'add-transaction'   as Page },
            { label: '📊 Reports',             page: 'reports'           as Page },
          ].map(a => (
            <button key={a.label} onClick={() => onNav(a.page)}
              className="bg-blue-600 hover:bg-orange-500 text-white rounded-xl py-3.5 px-4 text-sm font-semibold transition-all shadow-sm hover:shadow-md active:scale-[0.97] text-left lg:text-center">
              {a.label}
            </button>
          ))}
        </div>
      </div>

      {/* 6 Stat cards */}
      <div className="grid grid-cols-2 lg:grid-cols-4 xl:grid-cols-2 gap-3">
        <StatCard label="Today's Earnings"  value={fmt(todayTotal)} sub="All collected today"   emoji="💰" bgColor="bg-orange-50" />
        <StatCard label="Cash Collection"   value={fmt(todayCash)}  sub="Cash today"            emoji="💵" bgColor="bg-green-50"  />
        <StatCard label="UPI Collection"    value={fmt(todayUPI)}   sub="UPI today"             emoji="📱" bgColor="bg-blue-50"   />
        <StatCard label="Total Salary"  value={fmt(dashboard.totalSalary)} sub="Monthly Salary" emoji="🧾" bgColor="bg-orange-50" />
        <StatCard label="Pending Salary"    value={fmt(pendingSal)} sub="Total to pay workers"  emoji="💼" bgColor="bg-blue-50"   />
        <StatCard label="Worker Transaction"  value={fmt(dashboard.totalTransactions)} sub="Total Advances & Expenses" emoji="📋" bgColor="bg-purple-50"
/>
      </div>

      {/* Activity + Workers */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
        <div className="lg:col-span-2 bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <h3 className="font-semibold text-gray-900 mb-4">Recent Activity</h3>
          <div className="space-y-0.5">
            {recentActivity.map((item, i) => (
              <div key={i} className="flex items-center justify-between py-2.5 border-b border-gray-50 last:border-0">
                <div className="flex items-center gap-3 min-w-0">
                  <span className="text-xl flex-shrink-0">{item.emoji}</span>
                  <div className="min-w-0">
                    <p className="text-sm font-medium text-gray-800 truncate">{item.label}</p>
                    <p className="text-xs text-gray-400">{item.date}</p>
                  </div>
                </div>
                <span className={`text-sm font-semibold flex-shrink-0 ml-3 ${item.positive ? 'text-green-600' : 'text-gray-700'}`}>{item.amount}</span>
              </div>
            ))}
          </div>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <h3 className="font-semibold text-gray-900 mb-4">Workers</h3>
          <div className="space-y-3">
            {workers.map(w => {
              const rem = calcRemaining(w, transactions)
              const pct = Math.round(((w.salary - rem) / w.salary) * 100)
              return (
                <button key={w.id} onClick={() => onSelectWorker(w)} className="w-full text-left hover:bg-gray-50 rounded-xl p-2 transition-colors">
                  <div className="flex items-center gap-2.5 mb-1.5">
                    <div className="w-7 h-7 rounded-full text-white text-xs font-bold flex items-center justify-center flex-shrink-0"
                      style={{ backgroundColor: w.color }}>{w.initials}</div>
                    <p className="flex-1 text-sm font-medium text-gray-800 truncate">{w.name}</p>
                    <span className={`text-xs font-semibold ${rem > 0 ? 'text-orange-600' : 'text-green-600'}`}>{fmt(rem)}</span>
                  </div>
                  <div className="h-1 bg-gray-100 rounded-full ml-9 overflow-hidden">
                    <div className="h-full bg-blue-600 rounded-full" style={{ width: `${pct}%` }} />
                  </div>
                </button>
              )
            })}
          </div>
          <button onClick={() => onNav('workers')} className="mt-3 w-full text-blue-600 text-xs font-semibold hover:underline">
            View all workers →
          </button>
        </div>
      </div>

      {/* Pending payments alert */}
      {pending.length > 0 && (
        <button onClick={() => onNav('pending-payments')}
          className="w-full flex items-center justify-between bg-orange-50 border border-orange-200 rounded-xl px-5 py-3.5 hover:bg-orange-100 transition-colors">
          <div className="flex items-center gap-3">
            <span className="text-xl">🧾</span>
            <div className="text-left">
              <p className="font-semibold text-orange-800 text-sm">{pending.length} pending customer payment{pending.length > 1 ? 's' : ''}</p>
              <p className="text-xs text-orange-600">Total: {fmt(pendingAmt)} · Click to manage</p>
            </div>
          </div>
          <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2.5" className="text-orange-500">
            <path strokeLinecap="round" strokeLinejoin="round" d="M9 5l7 7-7 7" />
          </svg>
        </button>
      )}
    </div>
  )
}

// ─── Daily Earnings ───────────────────────────────────────────────────────────

function DailyEarningsPage({ earnings, setEarnings }: { earnings: Earning[]; setEarnings: React.Dispatch<React.SetStateAction<Earning[]>> }) {
  const [form, setForm] = useState({
    date: todayStr, vehicleType: 'Car', washType: 'Full Wash',
    amount: '', paymentMode: 'Cash', customerName: '', remarks: '',
  })
  const [saved, setSaved] = useState(false)
  const [search, setSearch] = useState('')
  const [editingId, setEditingId] = useState<number | null>(null);

  const todayEntries = earnings.filter(e => e.date === todayStr && e.paymentMode !== 'Pending')
  const todayCash    = todayEntries.filter(e => e.paymentMode === 'Cash').reduce((s, e) => s + e.amount, 0)
  const todayUPI     = todayEntries.filter(e => e.paymentMode === 'UPI').reduce((s, e) => s + e.amount, 0)

  const filtered = earnings.filter(e =>
  !search ||
  (e.vehicleType ?? "").toLowerCase().includes(search.toLowerCase()) ||
  (e.customerName ?? "").toLowerCase().includes(search.toLowerCase())
);

  async function handleSave() {
  if (!form.amount) return;

  try {
    const response = await fetch(
  editingId
    ? `https://washing-station-production.up.railway.app/api/earnings/${editingId}`
    : "https://washing-station-production.up.railway.app/api/earnings",
  {
    method: editingId ? "PUT" : "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      date: form.date,
      customerName:
        form.paymentMode === "Pending" ? form.customerName : "",
      vehicleType: form.vehicleType,
      washType: form.washType,
      amount: Number(form.amount),
      paymentMode: form.paymentMode,
      status: "Active",
      remarks:
        form.paymentMode === "Pending" ? form.remarks : "",
    }),
  }
);

    if (!response.ok) {
      throw new Error("Failed to save earning");
    }

    const savedEarning = await response.json();

    if (editingId) {
  setEarnings(prev =>
    prev.map(e => (e.id === editingId ? savedEarning : e))
  );
} else {
  setEarnings(prev => [savedEarning, ...prev]);
}
setEditingId(null);

    setForm({
      date: todayStr,
      vehicleType: "Car",
      washType: "Full Wash",
      amount: "",
      paymentMode: "Cash",
      customerName: "",
      remarks: "",
    });

    setSaved(true);
    setTimeout(() => setSaved(false), 2500);

  } catch (err) {
    console.error(err);
    alert("Unable to save earning.");
  }
}

  const fc = 'w-full border border-gray-200 rounded-xl px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white'
  

  return (
    <div className="p-4 lg:p-6 space-y-5">
      {/* Form */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
        <h3 className="font-semibold text-gray-900 mb-4">Add Earning Entry</h3>
        {saved && <div className="bg-green-50 border border-green-200 text-green-700 rounded-xl px-4 py-2.5 text-sm font-medium mb-4">✅ Entry saved!</div>}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">Date</label>
            <input type="date" value={form.date} onChange={e => setForm(f => ({ ...f, date: e.target.value }))} className={fc} />
          </div>
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">Vehicle Type</label>
            <select value={form.vehicleType} onChange={e => setForm(f => ({ ...f, vehicleType: e.target.value }))} className={fc}>
              {['Bike', 'Car', 'TT','Tempo/Pickup','Tipper','Lorry','Auto'].map(v => <option key={v}>{v}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">Wash Type</label>
            <select value={form.washType} onChange={e => setForm(f => ({ ...f, washType: e.target.value }))} className={fc}>
              {['Body Wash','Water Wash','Full Wash','Greasing'].map(v => <option key={v}>{v}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">Amount (₹)</label>
            <input type="number" value={form.amount} onChange={e => setForm(f => ({ ...f, amount: e.target.value }))} placeholder="e.g. 450" className={fc} />
          </div>
          <div>
            <label className="text-xs font-medium text-gray-600 mb-2 block">Payment Mode</label>
            <div className="flex gap-4 flex-wrap mt-1">
              {['Cash', 'UPI', 'Pending'].map(mode => (
                <label key={mode} className="flex items-center gap-2 cursor-pointer">
                  <input type="radio" name="paymentMode" value={mode} checked={form.paymentMode === mode}
                    onChange={() => setForm(f => ({ ...f, paymentMode: mode }))} className="accent-blue-600 w-4 h-4" />
                  <span className={`text-sm font-medium ${mode === 'Pending' ? 'text-orange-600' : 'text-gray-700'}`}>{mode}</span>
                </label>
              ))}
            </div>
          </div>
          {form.paymentMode === 'Pending' && (
            <>
              <div>
                <label className="text-xs font-medium text-gray-600 mb-1.5 block">Customer Name</label>
                <input type="text" value={form.customerName} onChange={e => setForm(f => ({ ...f, customerName: e.target.value }))}
                  placeholder="e.g. Kiran Patel" className={fc} />
              </div>
              <div className="sm:col-span-2 lg:col-span-3">
                <label className="text-xs font-medium text-gray-600 mb-1.5 block">Remarks <span className="text-gray-400">(optional)</span></label>
                <input type="text" value={form.remarks} onChange={e => setForm(f => ({ ...f, remarks: e.target.value }))}
                  placeholder="e.g. Will pay tomorrow" className={fc} />
              </div>
            </>
          )}
        </div>
        <button onClick={handleSave} className="mt-5 bg-blue-600 hover:bg-blue-700 text-white font-semibold px-8 py-3 rounded-xl transition-colors shadow-sm">
          {editingId ? "Update Entry" : "Save Entry"}
        </button>
      </div>

      {/* Summary */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        {[
          { label: 'Cash Today',   value: fmt(todayCash),              color: 'text-green-600' },
          { label: 'UPI Today',    value: fmt(todayUPI),               color: 'text-blue-600' },
          { label: 'Total Earned', value: fmt(todayCash + todayUPI),   color: 'text-gray-900' },
          { label: 'Vehicles',     value: String(todayEntries.length), color: 'text-orange-600' },
        ].map(s => (
          <div key={s.label} className="bg-white rounded-xl border border-gray-100 shadow-sm p-4">
            <p className="text-xs text-gray-500 mb-1">{s.label}</p>
            <p className={`text-xl font-bold ${s.color}`}>{s.value}</p>
          </div>
        ))}
      </div>

      {/* Table */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
        <div className="flex items-center justify-between mb-4 gap-3">
          <h3 className="font-semibold text-gray-900">All Entries</h3>
          <input type="text" placeholder="Search..." value={search} onChange={e => setSearch(e.target.value)}
            className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-44" />
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm min-w-[500px]">
            <thead>
              <tr className="border-b border-gray-100">
                {['Date', 'Vehicle', 'Wash Type', 'Amount', 'Mode', 'Customer','Actions'].map(h => (
                  <th key={h} className="text-left py-2.5 px-3 text-xs font-semibold text-gray-500">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(e => (
                <tr key={e.id} className="border-b border-gray-50 hover:bg-gray-50 transition-colors">
                  <td className="py-2.5 px-3 text-gray-500 text-xs">{e.date}</td>
                  <td className="py-2.5 px-3 text-gray-700">{e.vehicleType}</td>
                  <td className="py-2.5 px-3 text-gray-600 text-xs">{e.washType}</td>
                  <td className="py-2.5 px-3 font-semibold text-gray-900">{fmt(e.amount)}</td>
                  <td className="py-2.5 px-3">
                    <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${
                      e.paymentMode === 'Cash' ? 'bg-green-100 text-green-700' :
                      e.paymentMode === 'UPI'  ? 'bg-blue-100 text-blue-700' :
                      'bg-orange-100 text-orange-700'
                    }`}>{e.paymentMode}</span>
                  </td>
                  <td className="py-2.5 px-3 text-gray-600">
  {e.customerName || "—"}
</td>

<td className="py-2.5 px-3">
  <div className="flex gap-2">
    <button
      onClick={() => {
        setEditingId(e.id);

        setForm({
          date: e.date,
          vehicleType: e.vehicleType,
          washType: e.washType,
          amount: String(e.amount),
          paymentMode: e.paymentMode,
          customerName: e.customerName || "",
          remarks: e.remarks || "",
        });
      }}
      className="px-4 py-2 rounded-lg bg-blue-600 text-white hover:bg-blue-700 text-sm font-medium"
    >
      Edit
    </button>

    <button
      onClick={async () => {
        if (!confirm("Delete this earning?")) return;

        await fetch(`https://washing-station-production.up.railway.app/api/earnings/${e.id}`, {
          method: "DELETE",
        });

        setEarnings(prev => prev.filter(x => x.id !== e.id));
      }}
      className="px-4 py-2 rounded-lg bg-red-600 text-white hover:bg-red-700 text-sm font-medium"
    >
      Delete
    </button>
  </div>
</td>

</tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={7} className="py-8 text-center text-gray-400 text-sm">No entries found.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

// ─── Pending Payments ─────────────────────────────────────────────────────────

function PendingPaymentsPage({ earnings, setEarnings }: { earnings: Earning[]; setEarnings: React.Dispatch<React.SetStateAction<Earning[]>> }) {
  const [tab, setTab]         = useState<'pending' | 'paid'>('pending')
  const [payModal, setPayModal] = useState<Earning | null>(null)
  const [payMode, setPayMode]  = useState<'Cash' | 'UPI'>('Cash')

  const pending = activePending(earnings)
  const paid    = earnings.filter(e => e.status === 'Paid')
  const pendingTotal = pending.reduce((s, e) => s + e.amount, 0)
  const paidTotal    = paid.reduce((s, e) => s + e.amount, 0)

  async function handleConfirm() {
  if (!payModal) return;

  try {
    const response = await fetch(
      `https://washing-station-production.up.railway.app/api/earnings/${payModal.id}`,
      {
        method: "PUT",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          ...payModal,
          paymentMode: payMode,
          status: "Paid",
        }),
      }
    );

    if (!response.ok) {
      throw new Error("Failed to update payment");
    }

    const updated = await response.json();

    setEarnings(prev =>
      prev.map(e =>
        e.id === updated.id
          ? {
              ...updated,
              paidDate: new Date().toISOString().split("T")[0],
              paidMode: payMode,
            }
          : e
      )
    );

    setPayModal(null);

  } catch (err) {
    console.error(err);
    alert("Unable to update payment.");
  }
}

  const tabBtn = (t: typeof tab, label: string, count?: number) => (
    <button onClick={() => setTab(t)}
      className={`flex items-center gap-1.5 px-5 py-2 rounded-xl text-sm font-semibold transition-all ${tab === t ? 'bg-blue-600 text-white shadow-sm' : 'text-gray-600 hover:bg-gray-100'}`}>
      {label}
      {count !== undefined && (
        <span className={`text-xs px-1.5 py-0.5 rounded-full ${tab === t ? 'bg-white/20 text-white' : 'bg-gray-200 text-gray-600'}`}>{count}</span>
      )}
    </button>
  )

  return (
    <div className="p-4 lg:p-6 space-y-5">
      {/* Summary */}
      <div className="grid grid-cols-2 gap-4">
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5 flex items-start justify-between">
          <div>
            <p className="text-xs text-gray-500 mb-1">Active Pending</p>
            <p className="text-2xl font-bold text-orange-600">{fmt(pendingTotal)}</p>
            <p className="text-xs text-gray-400 mt-1">{pending.length} customers</p>
          </div>
          <div className="w-10 h-10 bg-orange-50 rounded-xl flex items-center justify-center text-xl">🧾</div>
        </div>
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5 flex items-start justify-between">
          <div>
            <p className="text-xs text-gray-500 mb-1">Cleared Payments</p>
            <p className="text-2xl font-bold text-green-600">{fmt(paidTotal)}</p>
            <p className="text-xs text-gray-400 mt-1">{paid.length} payments received</p>
          </div>
          <div className="w-10 h-10 bg-green-50 rounded-xl flex items-center justify-center text-xl">✅</div>
        </div>
      </div>

      {/* Tabs */}
      <div className="flex gap-2 bg-gray-100 p-1 rounded-xl w-fit">
        {tabBtn('pending', '⏳ Pending', pending.length)}
        {tabBtn('paid', '✅ Paid', paid.length)}
      </div>

      {/* Pending table */}
      {tab === 'pending' && (
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <h3 className="font-semibold text-gray-900 mb-4">Pending Customer Payments</h3>
          {pending.length === 0 ? (
            <div className="text-center py-10">
              <p className="text-3xl mb-2">✅</p>
              <p className="font-semibold text-gray-700">All payments cleared!</p>
              <p className="text-sm text-gray-400 mt-1">No pending payments at the moment.</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm min-w-[560px]">
                <thead>
                  <tr className="border-b border-gray-100">
                    {['Customer Name', 'Vehicle', 'Amount', 'Pending Date', 'Remarks', 'Action'].map(h => (
                      <th key={h} className="text-left py-2.5 px-3 text-xs font-semibold text-gray-500">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {pending.map(e => (
                    <tr key={e.id} className="border-b border-gray-50 hover:bg-gray-50">
                      <td className="py-3 px-3">
                        <p className="font-medium text-gray-800">{e.customerName || 'Walk-in'}</p>
                      </td>
                      <td className="py-3 px-3 text-gray-600">{e.vehicleType} · {e.washType}</td>
                      <td className="py-3 px-3 font-bold text-orange-600">{fmt(e.amount)}</td>
                      <td className="py-3 px-3 text-gray-500 text-xs">{e.date}</td>
                      <td className="py-3 px-3 text-gray-500 text-xs">{e.remarks || '—'}</td>
                      <td className="py-3 px-3">
                        <button onClick={() => { setPayModal(e); setPayMode('Cash') }}
                          className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold px-3 py-1.5 rounded-lg transition-colors whitespace-nowrap">
                          ✔ Mark as Paid
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* Paid table */}
      {tab === 'paid' && (
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <h3 className="font-semibold text-gray-900 mb-4">Cleared Payments</h3>
          {paid.length === 0 ? (
            <p className="text-gray-400 text-sm text-center py-8">No payments cleared yet.</p>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-sm min-w-[500px]">
                <thead>
                  <tr className="border-b border-gray-100">
                    {['Customer', 'Vehicle', 'Amount', 'Pending Date', 'Paid Date', 'Mode'].map(h => (
                      <th key={h} className="text-left py-2.5 px-3 text-xs font-semibold text-gray-500">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {paid.map(e => (
                    <tr key={e.id} className="border-b border-gray-50 hover:bg-gray-50">
                      <td className="py-3 px-3 font-medium text-gray-800">{e.customerName}</td>
                      <td className="py-3 px-3 text-gray-600 text-xs">{e.vehicleType}</td>
                      <td className="py-3 px-3 font-bold text-green-600">{fmt(e.amount)}</td>
                      <td className="py-3 px-3 text-gray-500 text-xs">{e.date}</td>
                      <td className="py-3 px-3 text-gray-500 text-xs">{e.paidDate}</td>
                      <td className="py-3 px-3">
                        <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${e.paidMode === 'Cash' ? 'bg-green-100 text-green-700' : 'bg-blue-100 text-blue-700'}`}>
                          {e.paidMode}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* Payment modal */}
      {payModal && (
        <>
          <div className="fixed inset-0 bg-black/40 z-40" onClick={() => setPayModal(null)} />
          <div className="fixed inset-0 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-2xl shadow-2xl max-w-sm w-full p-6">
              <h3 className="font-bold text-gray-900 text-lg mb-1">Mark as Paid</h3>
              <p className="text-sm text-gray-500 mb-5">Confirm payment from <strong className="text-gray-800">{payModal.customerName}</strong></p>
              <div className="bg-blue-50 rounded-xl p-3 mb-5 text-center">
                <p className="text-xs text-blue-600 mb-0.5">Amount</p>
                <p className="text-2xl font-bold text-blue-700">{fmt(payModal.amount)}</p>
              </div>
              <div className="mb-5">
                <p className="text-xs font-medium text-gray-600 mb-2">Payment Mode</p>
                <div className="flex gap-4">
                  {(['Cash', 'UPI'] as const).map(m => (
                    <label key={m} className="flex items-center gap-2 cursor-pointer">
                      <input type="radio" name="payMode" value={m} checked={payMode === m}
                        onChange={() => setPayMode(m)} className="accent-blue-600 w-4 h-4" />
                      <span className="text-sm font-medium text-gray-700">{m}</span>
                    </label>
                  ))}
                </div>
              </div>
              <div className="flex gap-3">
                <button onClick={() => setPayModal(null)}
                  className="flex-1 border border-gray-200 text-gray-700 font-semibold py-3 rounded-xl hover:bg-gray-50 transition-colors text-sm">Cancel</button>
                <button onClick={handleConfirm}
                  className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 rounded-xl transition-colors shadow-sm text-sm">
                  Confirm Payment
                </button>
              </div>
            </div>
          </div>
        </>
      )}
    </div>
  )
}

// ─── Workers ──────────────────────────────────────────────────────────────────

interface WorkersPageProps {
  workers: Worker[]
  setWorkers: React.Dispatch<React.SetStateAction<Worker[]>>
  transactions: WTransaction[]
  onSelectWorker: (w: Worker) => void
}

function WorkersPage({ workers, setWorkers, transactions, onSelectWorker }: WorkersPageProps) {
  const [showForm, setShowForm] = useState(false)
  const [form, setForm] = useState({ name: '', salary: '', joinDate: todayStr })
  const COLORS = ['#2563EB','#F97316','#10B981','#8B5CF6','#EC4899','#14B8A6']

  async function handleAdd() {
  if (!form.name || !form.salary) return;

  try {
    const response = await fetch("https://washing-station-production.up.railway.app/api/workers", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        name: form.name,
        salary: Number(form.salary),
        joinDate: form.joinDate,
        status: "Active",
      }),
    });

    if (!response.ok) {
      throw new Error("Failed to save worker");
    }

    const savedWorker = await response.json();

    const initials = savedWorker.name
      .split(" ")
      .map((n: string) => n[0])
      .join("")
      .toUpperCase()
      .slice(0, 2);

    const COLORS = [
      "#2563EB",
      "#F97316",
      "#10B981",
      "#8B5CF6",
      "#EC4899",
      "#14B8A6",
    ];

    setWorkers(prev => [
      ...prev,
      {
        ...savedWorker,
        initials,
        color: COLORS[prev.length % COLORS.length],
      },
    ]);

    setForm({
      name: "",
      salary: "",
      joinDate: todayStr,
    });

    setShowForm(false);

  } catch (err) {
    console.error(err);
    alert("Unable to save worker.");
  }
}
 async function toggleWorkerStatus(id: number) {
  const worker = workers.find(w => w.id === id);

  if (!worker) return;

  const updatedWorker = {
    ...worker,
    status: worker.status === "Active" ? "Inactive" : "Active",
  };

  try {
    const response = await fetch(
      `https://washing-station-production.up.railway.app/api/workers/${id}`,
      {
        method: "PUT",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(updatedWorker),
      }
    );

    if (!response.ok) {
      throw new Error("Failed to update worker");
    }

    const savedWorker = await response.json();

    setWorkers(prev =>
      prev.map(w => (w.id === id ? { ...w, ...savedWorker } : w))
    );

  } catch (error) {
    console.error(error);
    alert("Unable to update worker.");
  }
}

async function removeWorker(id: number) {
  const worker = workers.find(w => w.id === id);

  if (!worker) return;

  if (worker.status === "Active") {
    alert("Please deactivate the worker before removing.");
    return;
  }

  const confirmed = window.confirm(
    `Are you sure you want to permanently remove ${worker.name}?`
  );

  if (!confirmed) return;

  try {
    const response = await fetch(
      `https://washing-station-production.up.railway.app/api/workers/${id}`,
      {
        method: "DELETE",
      }
    );

    if (!response.ok) {
      throw new Error("Failed to delete worker");
    }

    setWorkers(prev => prev.filter(w => w.id !== id));

  } catch (err) {
    console.error(err);
    alert("Unable to delete worker.");
  }
}
  const fc = 'w-full border border-gray-200 rounded-xl px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white'

  return (
    <div className="p-4 lg:p-6 space-y-5">
      <div className="flex items-center justify-between">
        <p className="text-sm text-gray-500">
  {
    workers.filter(w => w.status === "Active").length
  } active workers
</p>
        <button onClick={() => setShowForm(v => !v)}
          className="bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold px-4 py-2.5 rounded-xl transition-colors">
          + Add Worker
        </button>
      </div>

      {showForm && (
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <h3 className="font-semibold text-gray-900 mb-4">New Worker</h3>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            <div>
              <label className="text-xs font-medium text-gray-600 mb-1.5 block">Full Name</label>
              <input type="text" value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} placeholder="e.g. Ravi Kumar" className={fc} />
            </div>
            <div>
              <label className="text-xs font-medium text-gray-600 mb-1.5 block">Monthly Salary (₹)</label>
              <input type="number" value={form.salary} onChange={e => setForm(f => ({ ...f, salary: e.target.value }))} placeholder="12000" className={fc} />
            </div>
            <div>
              <label className="text-xs font-medium text-gray-600 mb-1.5 block">Joining Date</label>
              <input type="date" value={form.joinDate} onChange={e => setForm(f => ({ ...f, joinDate: e.target.value }))} className={fc} />
            </div>
          </div>
          <div className="flex gap-3 mt-4">
            <button onClick={handleAdd} className="bg-blue-600 hover:bg-blue-700 text-white font-semibold px-6 py-2.5 rounded-xl transition-colors">Save Worker</button>
            <button onClick={() => setShowForm(false)} className="text-gray-600 hover:text-gray-900 font-medium px-4 py-2.5 rounded-xl hover:bg-gray-100 transition-colors">Cancel</button>
          </div>
        </div>
      )}

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-gray-100 bg-gray-50">
              {['Worker', 'Monthly Salary', 'Joining Date', 'Remaining Salary', 'Status', 'Action'].map(h => (
                <th key={h} className="text-left py-3 px-4 text-xs font-semibold text-gray-500">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {workers.map(w => {
              const rem = calcRemaining(w, transactions)
              return (
                <tr
  key={w.id}
  className={`border-b border-gray-50 transition-colors ${
    w.status === "Inactive"
      ? "bg-gray-100 opacity-70"
      : "hover:bg-gray-50"
  }`}
>
                  <td className="py-3.5 px-4">
                    <div className="flex items-center gap-2.5">
                      <div className="w-8 h-8 rounded-full text-white text-xs font-bold flex items-center justify-center flex-shrink-0"
                        style={{ backgroundColor: w.color }}>{w.initials}</div>
                      <p className="font-medium text-gray-800">{w.name}</p>
                    </div>
                  </td>
                  <td className="py-3.5 px-4 font-semibold text-gray-900">{fmt(w.salary)}</td>
                  <td className="py-3.5 px-4 text-gray-500 text-xs">{w.joinDate}</td>
                  <td className="py-3.5 px-4">
                    <span className={`font-bold ${rem > 0 ? 'text-orange-600' : 'text-green-600'}`}>{fmt(rem)}</span>
                  </td>
                  <td className="py-3.5 px-4">
  <span
    className={`px-2.5 py-0.5 rounded-full text-xs font-semibold ${
      w.status === "Active"
        ? "bg-green-100 text-green-700"
        : "bg-red-100 text-red-700"
    }`}
  >
    {w.status}
  </span>
</td>
                  <td className="py-3.5 px-4">
  <div className="flex gap-2 flex-wrap">

    <button
      onClick={() => onSelectWorker(w)}
      className="bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-semibold py-1.5 px-3 rounded-lg transition-colors"
    >
      View
    </button>

    <button
      onClick={() => toggleWorkerStatus(w.id)}
      className={`text-xs font-semibold py-1.5 px-3 rounded-lg transition-colors ${
        w.status === "Active"
          ? "bg-orange-100 hover:bg-orange-200 text-orange-700"
          : "bg-green-100 hover:bg-green-200 text-green-700"
      }`}
    >
      {w.status === "Active"
        ? "Deactivate"
        : "Activate"}
    </button>

    <button
      onClick={() => removeWorker(w.id)}
      disabled={w.status === "Active"}
      className={`text-xs font-semibold py-1.5 px-3 rounded-lg transition-colors ${
        w.status === "Inactive"
          ? "bg-red-100 hover:bg-red-200 text-red-700"
          : "bg-gray-100 text-gray-400 cursor-not-allowed"
      }`}
    >
      Remove
    </button>

  </div>
</td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ─── Worker Detail ────────────────────────────────────────────────────────────

function WorkerDetailPage({ worker, transactions, onBack }: { worker: Worker; transactions: WTransaction[]; onBack: () => void }) {
  const [txs, setTxs] = useState<WTransaction[]>([]);
  useEffect(() => {
    fetch(`https://washing-station-production.up.railway.app/api/transactions/worker/${worker.id}`)
        .then(res => res.json())
        .then(data => {

            console.log("Backend Response:", data);

            const mapped = data.map((tx: any) => ({
                id: tx.id,
                workerId: tx.worker.id,
                workerName: tx.worker.name,
                date: tx.date,
                type: tx.type,
                amount: tx.amount,
                remarks: tx.remarks
            }));

            console.log("Mapped:", mapped);

            setTxs(mapped);
        })
        .catch(console.error);
}, [worker.id]);
  const deductions = txs.filter(t => t.type !== 'Salary Settlement').reduce((s, t) => s + t.amount, 0)
  const salaryPaid = txs.filter(t => t.type === 'Salary Settlement').reduce((s, t) => s + t.amount, 0)
  const remaining  = Math.max(0, worker.salary - deductions - salaryPaid)

  return (
    <div className="p-4 lg:p-6 space-y-5">
      <button onClick={onBack} className="flex items-center gap-1.5 text-blue-600 hover:text-blue-700 text-sm font-medium">
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2.5"><path strokeLinecap="round" strokeLinejoin="round" d="M15 19l-7-7 7-7" /></svg>
        Back to Workers
      </button>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5 flex flex-col sm:flex-row items-start sm:items-center gap-4">
        <div className="w-16 h-16 rounded-full flex items-center justify-center text-white font-bold text-xl flex-shrink-0"
          style={{ backgroundColor: worker.color }}>{worker.initials}</div>
        <div className="flex-1">
          <h2 className="text-xl font-bold text-gray-900">{worker.name}</h2>
          <p className="text-gray-500 text-sm">Joined {worker.joinDate}</p>
        </div>
        <div className="text-right">
          <p className="text-xs text-gray-500">Monthly Salary</p>
          <p className="text-2xl font-bold text-blue-600">{fmt(worker.salary)}</p>
          <p className={`text-sm font-semibold ${remaining > 0 ? 'text-orange-600' : 'text-green-600'}`}>
            {remaining > 0 ? `${fmt(remaining)} remaining` : '✅ Fully paid'}
          </p>
        </div>
      </div>

      <div className="grid grid-cols-3 gap-4">
        <div className="bg-red-50 border border-red-100 rounded-xl p-4">
          <p className="text-xs text-red-600 font-medium mb-1">Total Deductions</p>
          <p className="text-xl font-bold text-red-700">{fmt(deductions)}</p>
        </div>
        <div className="bg-green-50 border border-green-100 rounded-xl p-4">
          <p className="text-xs text-green-600 font-medium mb-1">Salary Paid</p>
          <p className="text-xl font-bold text-green-700">{fmt(salaryPaid)}</p>
        </div>
        <div className="bg-orange-50 border border-orange-100 rounded-xl p-4">
          <p className="text-xs text-orange-600 font-medium mb-1">Remaining</p>
          <p className="text-xl font-bold text-orange-700">{fmt(remaining)}</p>
        </div>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
        <h3 className="font-semibold text-gray-900 mb-5">Transaction History</h3>
        {txs.length === 0 ? (
          <p className="text-gray-400 text-sm text-center py-8">No transactions yet.</p>
        ) : (
          <div className="relative pl-7 space-y-4">
            <div className="absolute left-2.5 top-2 bottom-2 w-px bg-gray-200" />
            {txs.map(tx => (
              <div key={tx.id} className="relative">
                <div className="absolute -left-[19px] top-3 w-3.5 h-3.5 rounded-full bg-white border-2 border-blue-400" />
                <div className="bg-gray-50 rounded-xl p-4">
                  <div className="flex items-center justify-between mb-1 gap-2">
                    <span className={`text-xs font-semibold px-2.5 py-0.5 rounded-full ${txBadgeColor(tx.type)}`}>{tx.type}</span>
                    <span className="text-sm font-bold text-gray-900">{fmt(tx.amount)}</span>
                  </div>
                  <p className="text-xs text-gray-500">{tx.date}{tx.remarks ? ` · ${tx.remarks}` : ''}</p>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  )
}

// ─── Add Transaction ──────────────────────────────────────────────────────────

interface AddTxProps {
  workers: Worker[]
  transactions: WTransaction[]
  setTransactions: React.Dispatch<React.SetStateAction<WTransaction[]>>
}

function AddTransactionPage({ workers, setTransactions }: AddTxProps) {
  const [form, setForm] = useState({
    workerName: workers[0]?.name || '',
    date: todayStr, type: 'Advance', amount: '', remarks: '',
  })
  const [saved, setSaved] = useState(false)

  const TX_TYPES = ['Advance', 'Rent', 'Bills', 'Food', 'Medical', 'Salary Settlement', 'Other']
  const fc = 'w-full border border-gray-200 rounded-xl px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white'

  async function handleSave() {
  if (!form.amount) return;

  const worker = workers.find(w => w.name === form.workerName);

  if (!worker) return;

  try {
    const response = await fetch("https://washing-station-production.up.railway.app/api/transactions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: form.type,
        amount: Number(form.amount),
        date: form.date,
        remarks: form.remarks,
        worker: {
          id: worker.id,
        },
      }),
    });

    if (!response.ok) {
      throw new Error("Failed to save transaction");
    }

    const savedTransaction = await response.json();

    setTransactions(prev => [
      ...prev,
      {
        id: savedTransaction.id,
        workerId: worker.id,
        workerName: worker.name,
        type: savedTransaction.type,
        amount: savedTransaction.amount,
        date: savedTransaction.date,
        remarks: savedTransaction.remarks,
      },
    ]);

    setForm(f => ({
      ...f,
      amount: "",
      remarks: "",
    }));

    setSaved(true);

    setTimeout(() => setSaved(false), 3000);

  } catch (error) {
    console.error(error);
    alert("Unable to save transaction.");
  }
}

  return (
    <div className="p-4 lg:p-6">
      <div className="max-w-lg mx-auto bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
        <h3 className="font-semibold text-gray-900 mb-6 text-lg">New Worker Transaction</h3>
        {saved && (
          <div className="bg-green-50 border border-green-200 text-green-700 rounded-xl px-4 py-3 text-sm font-medium mb-5">
            ✅ Transaction saved! Remaining salary updated automatically.
          </div>
        )}
        <div className="space-y-4">
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">Worker</label>
            <select value={form.workerName} onChange={e => setForm(f => ({ ...f, workerName: e.target.value }))} className={fc}>
              {workers.map(w => <option key={w.id} value={w.name}>{w.name}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">Date</label>
            <input type="date" value={form.date} onChange={e => setForm(f => ({ ...f, date: e.target.value }))} className={fc} />
          </div>
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">Transaction Type</label>
            <select value={form.type} onChange={e => setForm(f => ({ ...f, type: e.target.value }))} className={fc}>
              {TX_TYPES.map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">Amount (₹)</label>
            <input type="number" value={form.amount} onChange={e => setForm(f => ({ ...f, amount: e.target.value }))} placeholder="e.g. 2000" className={fc} />
          </div>
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">Remarks <span className="text-gray-400">(optional)</span></label>
            <textarea value={form.remarks} onChange={e => setForm(f => ({ ...f, remarks: e.target.value }))}
              placeholder="Optional note..." rows={3} className={`${fc} resize-none`} />
          </div>
          <button onClick={handleSave} className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 rounded-xl transition-colors shadow-sm">
            Save Transaction
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── Salary Settlement ────────────────────────────────────────────────────────

interface SalarySettlementProps {
  workers: Worker[]
  transactions: WTransaction[]
  settlements: SettlementRecord[]
  setSettlements: React.Dispatch<React.SetStateAction<SettlementRecord[]>>
  loadTransactions: () => void;
}

function SalarySettlementPage({ workers, transactions, settlements, setSettlements, loadTransactions }: SalarySettlementProps) {
  const [tab, setTab]         = useState<'current' | 'history' | 'reports'>('current')
  const today = new Date();

const [month, setMonth] = useState(today.getMonth() + 1);
const [year, setYear] = useState(today.getFullYear());
  const [search, setSearch]   = useState('')
  const [histSearch, setHistSearch] = useState('')
  const [selWorker, setSelWorker]   = useState<Worker | null>(null)
  const [showConfirm, setShowConfirm] = useState(false)
  const [showSuccess, setShowSuccess] = useState(false)
  console.log("Salary Page settlements:", settlements);

  function isSettled(wid: number) {
  return settlements.some(
    s =>
      s.workerId === wid &&
      s.month === month &&
      s.year === year &&
      s.settled
  );
}

  async function handleSettle() {
  if (!selWorker) return;

  try {
    const response = await fetch(
      `https://washing-station-production.up.railway.app/api/salary-settlements/${selWorker.id}`,
      {
        method: "POST",
      }
    );

    if (!response.ok) {
      const message = await response.text();
      if (response.status === 409) {
    alert("Salary has already been settled for this month.");
    return;
  }

  throw new Error(errorText);
    }

    const settlement = await response.json();

    setSettlements(prev => [
      ...prev.filter(
        s =>
          !(
            s.workerId === settlement.worker.id &&
            s.month === settlement.month &&
            s.year === settlement.year
          )
      ),
      {
        workerId: settlement.worker.id,
        month: settlement.month,
        year: settlement.year,
        settled: true,
        settlementDate: settlement.settlementDate,
      },
    ]);
    loadTransactions();
    setShowConfirm(false);
    setShowSuccess(true);

    setTimeout(() => {
      setShowSuccess(false);
      setSelWorker(null);
    }, 2500);

  } catch (err: any) {
  console.error(err);

  if (err.message.includes("409")) {
    alert("Salary has already been settled for this month.");
  } else {
    alert("Unable to settle salary.");
  }
}
}

  const filtered = workers.filter(w => !search || w.name.toLowerCase().includes(search.toLowerCase()))

  const totals = workers.reduce((acc, w) => {
    acc.salary     += w.salary
    acc.deductions += calcDeductions(w, transactions)
    acc.remaining  += calcRemaining(w, transactions)
    return acc
  }, { salary: 0, deductions: 0, remaining: 0 })
  const totalWorkers = workers.length;
  const settledWorkers = workers.filter(w => isSettled(w.id)).length;
  const pendingWorkers = totalWorkers - settledWorkers;
  const totalSalary = workers.reduce(
    (sum, w) => sum + w.salary,
    0
  );

  const totalDeductions = workers.reduce(
    (sum, w) => sum + calcDeductions(w, transactions),
    0
  );
  const totalSalaryPaid = workers.reduce((sum, w) => {
    if (!isSettled(w.id)) return sum;
    return sum + (w.salary - calcDeductions(w, transactions));
  }, 0);

  const histRows = settlements
    .filter(s => s.settled)
    .filter(s => {
      if (!histSearch) return true
      const w = workers.find(w => w.id === s.workerId)
      return w?.name.toLowerCase().includes(histSearch.toLowerCase()) ||
        MONTHS[s.month - 1].toLowerCase().includes(histSearch.toLowerCase())
    })
    .sort((a, b) => b.year - a.year || b.month - a.month)

  const tabBtn = (t: typeof tab, label: string) => (
    <button onClick={() => setTab(t)}
      className={`px-5 py-2 rounded-xl text-sm font-semibold transition-all ${tab === t ? 'bg-blue-600 text-white shadow-sm' : 'text-gray-600 hover:bg-gray-100'}`}>
      {label}
    </button>
  )

  const card = 'bg-white rounded-2xl shadow-sm border border-gray-100 p-5'

  return (
    <div className="p-4 lg:p-6 space-y-5">
      <div className="flex gap-2 bg-gray-100 p-1 rounded-xl w-fit">
        {tabBtn('current', '📋 Current Month')}
        {tabBtn('history', '🕐 History')}
        {tabBtn('reports', '📊 Reports')}
      </div>

      {tab === 'current' && (
        <>
          {/* Summary */}
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            {[
              { label: 'Total Workers',    value: String(workers.length), bg: 'bg-blue-50',   emoji: '👷' },
              { label: 'Total Salary',     value: fmt(totals.salary),     bg: 'bg-orange-50', emoji: '💼' },
              { label: 'Total Deductions', value: fmt(totals.deductions), bg: 'bg-red-50',    emoji: '📉' },
              { label: 'Pending Salary',   value: fmt(totals.remaining),  bg: 'bg-yellow-50', emoji: '⏳' },
            ].map(c => (
              <div key={c.label} className={`${card} flex items-start justify-between gap-2`}>
                <div>
                  <p className="text-xs text-gray-500 mb-1">{c.label}</p>
                  <p className="text-xl font-bold text-gray-900">{c.value}</p>
                </div>
                <div className={`w-9 h-9 rounded-xl flex items-center justify-center text-lg ${c.bg} flex-shrink-0`}>{c.emoji}</div>
              </div>
            ))}
          </div>

          {/* Filters */}
          <div className={card}>
            <div className="flex flex-wrap gap-3 items-end">
              <div>
                <label className="text-xs font-medium text-gray-600 mb-1.5 block">Month</label>
                <select value={month} onChange={e => setMonth(Number(e.target.value))}
                  className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                  {MONTHS.map((m, i) => <option key={m} value={i + 1}>{m}</option>)}
                </select>
              </div>
              <div>
                <label className="text-xs font-medium text-gray-600 mb-1.5 block">Year</label>
                <select value={year} onChange={e => setYear(Number(e.target.value))}
                  className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
                  {[2024, 2025, 2026].map(y => <option key={y}>{y}</option>)}
                </select>
              </div>
              <div className="flex-1 min-w-[160px]">
                <label className="text-xs font-medium text-gray-600 mb-1.5 block">Search Worker</label>
                <input type="text" value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by name..."
                  className="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white" />
              </div>
            </div>
          </div>

          {/* Table */}
          <div className={card}>
            <div className="overflow-x-auto">
              <table className="w-full text-sm min-w-[700px]">
                <thead>
                  <tr className="border-b border-gray-100 bg-gray-50">
                    {['Worker', 'Monthly Salary', 'Total Deductions', 'Remaining Salary', 'Status', 'Action'].map(h => (
                      <th key={h} className="text-left py-3 px-3 text-xs font-semibold text-gray-500">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map(w => {
                    const deductions = calcDeductions(w, transactions)
                    const remaining  = calcRemaining(w, transactions)
                    const settled    = isSettled(w.id)
                    return (
                      <tr key={w.id} className="border-b border-gray-50 hover:bg-gray-50">
                        <td className="py-3 px-3">
                          <div className="flex items-center gap-2.5">
                            <div className="w-8 h-8 rounded-full text-white text-xs font-bold flex items-center justify-center"
                              style={{ backgroundColor: w.color }}>{w.initials}</div>
                            <p className="font-medium text-gray-800">{w.name}</p>
                          </div>
                        </td>
                        <td className="py-3 px-3 font-semibold text-gray-900">{fmt(w.salary)}</td>
                        <td className="py-3 px-3 text-red-500 font-medium">{fmt(deductions)}</td>
                        <td className="py-3 px-3">
                          <span className={`font-bold ${settled ? 'text-green-600' : 'text-orange-600'}`}>
                            {settled ? '₹0' : fmt(remaining)}
                          </span>
                        </td>
                        <td className="py-3 px-3">
                          <span className={`px-2.5 py-0.5 rounded-full text-xs font-semibold ${settled ? 'bg-green-100 text-green-700' : 'bg-orange-100 text-orange-700'}`}>
                            {settled ? 'Settled' : 'Pending'}
                          </span>
                        </td>
                        <td className="py-3 px-3">
                          <button onClick={() => { setSelWorker(w); setShowSuccess(false) }}
                            className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold px-3 py-1.5 rounded-lg transition-colors">
                            View
                          </button>
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}

      {tab === 'history' && (
        <div className={card}>
          <div className="flex items-center justify-between mb-4 gap-3 flex-wrap">
            <h3 className="font-semibold text-gray-900">Settlement History</h3>
            <input type="text" value={histSearch} onChange={e => setHistSearch(e.target.value)}
              placeholder="Search worker or month..."
              className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-52 bg-white" />
          </div>
          <div className="space-y-3">
            {histRows.length === 0 && <p className="text-gray-400 text-sm text-center py-8">No settlement history found.</p>}
            {histRows.map((s, i) => {
              const w = workers.find(w => w.id === s.workerId)
              if (!w) return null
              const deductions = calcDeductions(w, transactions)
              const paid = w.salary - deductions;
              return (
                <div key={i} className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 p-4 bg-gray-50 rounded-xl hover:bg-gray-100 transition-colors">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-full text-white text-sm font-bold flex items-center justify-center flex-shrink-0"
                      style={{ backgroundColor: w.color }}>{w.initials}</div>
                    <div>
                      <p className="font-semibold text-gray-800">{w.name}</p>
                      <p className="text-xs text-gray-500">{MONTHS[s.month - 1]} {s.year}</p>
                    </div>
                  </div>
                  <div className="flex flex-wrap gap-4 text-sm">
                    <div><p className="text-xs text-gray-400">Salary</p><p className="font-semibold text-gray-800">{fmt(w.salary)}</p></div>
                    <div><p className="text-xs text-gray-400">Deductions</p><p className="font-semibold text-red-500">{fmt(deductions)}</p></div>
                    <div><p className="text-xs text-gray-400">Paid</p><p className="font-semibold text-green-600">{fmt(paid)}</p></div>
                    <div><p className="text-xs text-gray-400">Settlement Date</p><p className="font-semibold text-gray-700">{s.settlementDate}</p></div>
                  </div>
                  <span className="bg-green-100 text-green-700 px-3 py-1 rounded-full text-xs font-semibold">✅ Settled</span>
                </div>
              )
            })}
          </div>
        </div>
      )}

      {tab === 'reports' && (
        <div className="space-y-5">
          <div className={card}>
            <h3 className="font-semibold text-gray-900 mb-4">Per Worker Salary Report</h3>
            <p className="text-sm text-gray-500 mb-4">
              Payroll Month: {MONTHS[month - 1]} {year}
            </p>
            <div className="overflow-x-auto">
              <table className="w-full text-sm min-w-[600px]">
                <thead>
                  <tr className="border-b border-gray-100">
                    {['Worker', 'Salary', 'Deductions', 'Salary Paid', 'Remaining', 'Status'].map(h => (
                      <th key={h} className="text-left py-2.5 px-3 text-xs font-semibold text-gray-500">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {workers.map(w => {
                    const deductions = calcDeductions(w, transactions)
                    const remaining  = calcRemaining(w, transactions)
                    const paid = remaining === 0
                    ? w.salary - deductions
                    : 0;
                    return (
                      <tr key={w.id} className="border-b border-gray-50 hover:bg-gray-50">
                        <td className="py-3 px-3">
                          <div className="flex items-center gap-2">
                            <div className="w-6 h-6 rounded-full text-white text-xs font-bold flex items-center justify-center"
                              style={{ backgroundColor: w.color }}>{w.initials}</div>
                            <span className="font-medium text-gray-800">{w.name}</span>
                          </div>
                        </td>
                        <td className="py-3 px-3 text-gray-800">{fmt(w.salary)}</td>
                        <td className="py-3 px-3 text-red-500">{fmt(deductions)}</td>
                        <td className="py-3 px-3 text-green-600">{fmt(paid)}</td>
                        <td className="py-3 px-3 font-semibold text-orange-600">{fmt(remaining)}</td>
                        <td className="py-3 px-3">
                          <span className={`px-2.5 py-0.5 rounded-full text-xs font-semibold ${remaining === 0 ? 'bg-green-100 text-green-700' : 'bg-orange-100 text-orange-700'}`}>
                            {remaining === 0 ? 'Cleared' : 'Pending'}
                          </span>
                        </td>
                      </tr>
                    )
                  })}
                  <tr className="bg-blue-50 border-t-2 border-blue-200 font-bold">
                    <td className="py-3 px-3">TOTAL</td>
                    <td className="py-3 px-3">{fmt(totalSalary)}</td>
                    <td className="py-3 px-3 text-red-600">{fmt(totalDeductions)}</td>
                    <td className="py-3 px-3 text-green-600">{fmt(totalSalaryPaid)}</td>
                    <td className="py-3 px-3">{fmt(totals.remaining)}</td>
                    <td className="py-3 px-3">-</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Worker side panel */}
      {selWorker && (() => {
        const w   = selWorker
        const bd_deductions = calcDeductions(w, transactions)
        const bd_salaryPaid = transactions.filter(t => t.workerId === w.id && t.type === 'Salary Settlement').reduce((a, t) => a + t.amount, 0)
        const bd_remaining  = calcRemaining(w, transactions)
        const txs    = transactions.filter(t => t.workerId === w.id)
        const settled = isSettled(w.id)
        return (
          <>
            <div className="fixed inset-0 bg-black/40 z-40" onClick={() => { setSelWorker(null); setShowSuccess(false) }} />
            <div className="fixed right-0 top-0 h-full w-full max-w-md bg-white z-50 shadow-2xl flex flex-col overflow-hidden">
              <div className="flex items-center justify-between px-5 py-4 border-b border-gray-100">
                <h3 className="font-semibold text-gray-900">Worker Salary Details</h3>
                <button onClick={() => { setSelWorker(null); setShowSuccess(false) }}
                  className="p-1.5 rounded-lg hover:bg-gray-100 text-gray-500">
                  <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2.5"><path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" /></svg>
                </button>
              </div>
              <div className="flex-1 overflow-y-auto p-5 space-y-4">
                {showSuccess && (
                  <div className="bg-green-50 border border-green-200 rounded-xl p-4 text-center">
                    <p className="text-2xl mb-1">✅</p>
                    <p className="font-bold text-green-700">Salary Settled Successfully!</p>
                    <p className="text-xs text-green-600 mt-1">Status updated · Remaining set to ₹0</p>
                  </div>
                )}
                <div className="flex items-center gap-3">
                  <div className="w-14 h-14 rounded-full text-white font-bold text-lg flex items-center justify-center"
                    style={{ backgroundColor: w.color }}>{w.initials}</div>
                  <div>
                    <p className="font-bold text-gray-900">{w.name}</p>
                    <p className="text-xs text-gray-500">Joined {w.joinDate}</p>
                    <span className={`mt-1 inline-block px-2.5 py-0.5 rounded-full text-xs font-semibold ${settled ? 'bg-green-100 text-green-700' : 'bg-orange-100 text-orange-700'}`}>
                      {settled ? '✅ Settled' : '⏳ Pending'}
                    </span>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  {[
                    { label: 'Monthly Salary',  value: fmt(w.salary),       color: 'text-blue-600',   bg: 'bg-blue-50' },
                    { label: 'Total Deductions', value: fmt(bd_deductions),  color: 'text-red-500',    bg: 'bg-red-50' },
                    { label: 'Salary Paid',      value: fmt(bd_salaryPaid),  color: 'text-green-600',  bg: 'bg-green-50' },
                    { label: 'Remaining',        value: settled ? '₹0' : fmt(bd_remaining), color: settled ? 'text-green-600' : 'text-orange-600', bg: settled ? 'bg-green-50' : 'bg-orange-50' },
                  ].map(item => (
                    <div key={item.label} className={`${item.bg} rounded-xl p-3`}>
                      <p className="text-xs text-gray-500 mb-0.5">{item.label}</p>
                      <p className={`font-bold text-base ${item.color}`}>{item.value}</p>
                    </div>
                  ))}
                </div>
                <div>
                  <h4 className="font-semibold text-gray-800 text-sm mb-3">Transaction Timeline</h4>
                  {txs.length === 0
                    ? <p className="text-gray-400 text-xs text-center py-4">No transactions yet.</p>
                    : (
                      <div className="relative pl-6 space-y-3">
                        <div className="absolute left-2 top-1 bottom-1 w-px bg-gray-200" />
                        {txs.map(tx => (
                          <div key={tx.id} className="relative">
                            <div className="absolute -left-[17px] top-2.5 w-3 h-3 rounded-full bg-white border-2 border-blue-400" />
                            <div className="bg-gray-50 rounded-xl p-3">
                              <div className="flex items-center justify-between gap-2">
                                <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${txBadgeColor(tx.type)}`}>{tx.type}</span>
                                <span className="text-sm font-bold text-gray-900">{fmt(tx.amount)}</span>
                              </div>
                              <p className="text-xs text-gray-400 mt-1">{tx.date}{tx.remarks ? ` · ${tx.remarks}` : ''}</p>
                            </div>
                          </div>
                        ))}
                      </div>
                    )
                  }
                </div>
              </div>
              {!settled && !showSuccess && (
                <div className="p-5 border-t border-gray-100">
                  <button onClick={() => setShowConfirm(true)}
                    className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3.5 rounded-xl transition-colors shadow-sm text-sm">
                    💰 Settle Salary — {fmt(bd_remaining)}
                  </button>
                </div>
              )}
            </div>

            {showConfirm && (
              <>
                <div className="fixed inset-0 bg-black/60 z-60" />
                <div className="fixed inset-0 flex items-center justify-center z-70 p-4">
                  <div className="bg-white rounded-2xl shadow-2xl max-w-sm w-full p-6">
                    <div className="text-center mb-5">
                      <div className="w-14 h-14 bg-blue-50 rounded-full flex items-center justify-center text-2xl mx-auto mb-3">💰</div>
                      <h3 className="font-bold text-gray-900 text-lg mb-2">Confirm Salary Settlement</h3>
                      <p className="text-gray-500 text-sm leading-relaxed">
                        Are you sure you want to settle <strong>{w.name}</strong>'s salary? This will mark the month as fully paid while preserving all transaction history.
                      </p>
                    </div>
                    <div className="bg-blue-50 rounded-xl p-3 mb-5 text-center">
                      <p className="text-xs text-blue-600 mb-0.5">Amount to settle</p>
                      <p className="text-2xl font-bold text-blue-700">{fmt(bd_remaining)}</p>
                    </div>
                    <div className="flex gap-3">
                      <button onClick={() => setShowConfirm(false)}
                        className="flex-1 border border-gray-200 text-gray-700 font-semibold py-3 rounded-xl hover:bg-gray-50 text-sm">Cancel</button>
                      <button onClick={handleSettle}
                        className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 rounded-xl text-sm shadow-sm">Confirm Settlement</button>
                    </div>
                  </div>
                </div>
              </>
            )}
          </>
        )
      })()}
    </div>
  )
}

// ___ Expenses ________________________________________________________________

interface ExpensesPageProps {
  expenses: Expense[];
  setExpenses: React.Dispatch<React.SetStateAction<Expense[]>>;
}

function ExpensesPage({
  expenses,
  setExpenses,
}: ExpensesPageProps){
  console.log(expenses);
  const [date, setDate] = useState(
  new Date().toISOString().split("T")[0]
);

const [category, setCategory] = useState("Materials");

const [description, setDescription] = useState("");

const [amount, setAmount] = useState("");

const [paymentMode, setPaymentMode] = useState("Cash");

const [saved, setSaved] = useState(false);

const [editingId, setEditingId] = useState<number | null>(null);

const [search, setSearch] = useState("");

const [filterCategory, setFilterCategory] = useState("All");

const [filterPayment, setFilterPayment] = useState("All");

const expenseCategories = [
  "Materials",
  "Electricity",
  "Rent",
  "Water Bill",
  "Maintenance",
  "Fuel",
  "Other",
];

async function deleteExpense(id: number) {
  const confirmDelete = window.confirm(
    "Are you sure you want to delete this expense?"
  );

  if (!confirmDelete) return;

  try {
    const res = await fetch(`https://washing-station-production.up.railway.app/api/expenses/${id}`, {
      method: "DELETE",
    });

    if (!res.ok) {
      throw new Error("Delete failed");
    }

    setExpenses(prev => prev.filter(exp => exp.id !== id));
  } catch (err) {
    console.error(err);
    alert("Failed to delete expense.");
  }
}

function editExpense(exp: Expense) {
  setEditingId(exp.id);

  setDate(exp.date);
  setCategory(exp.category);
  setDescription(exp.description);
  setAmount(exp.amount.toString());
  setPaymentMode(exp.paymentMode);
}

async function addExpense() {
  if (!description.trim() || !amount) {
    alert("Please fill all required fields.");
    return;
  }

  const expense = {
    date,
    category,
    description,
    amount: Number(amount),
    paymentMode,
  };


    let res;

    if (editingId !== null) {
      res = await fetch(`https://washing-station-production.up.railway.app/api/expenses/${editingId}`, {
        method: "PUT",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(expense),
      });
    } else {
      res = await fetch("https://washing-station-production.up.railway.app/api/expenses", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(expense),
      });
    }
    const savedExpense = await res.json();

    if (editingId !== null) {
      setExpenses(prev =>
        prev.map(exp =>
          exp.id === editingId ? savedExpense : exp
        )
      );

      setEditingId(null);
    } else {
      setExpenses(prev => [...prev, savedExpense]);
    }

    setDescription("");
    setAmount("");
    setCategory("Materials");
    setPaymentMode("Cash");
  }

  const filteredExpenses = expenses.filter((exp) => {
  const matchesSearch =
    exp.description.toLowerCase().includes(search.toLowerCase());

  const matchesCategory =
    filterCategory === "All" || exp.category === filterCategory;

  const matchesPayment =
    filterPayment === "All" || exp.paymentMode === filterPayment;

  return matchesSearch && matchesCategory && matchesPayment;
  });

  return (
    <div className="p-4 lg:p-6">

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">

        <h2 className="text-xl font-bold mb-6">
          Expenses
        </h2>
        {saved && (
        <div className="bg-green-50 border border-green-200 text-green-700 rounded-xl px-4 py-3 mb-4 text-sm font-medium">
        ✅ Expense added successfully!
        </div>
        )}

      <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-4">

  <input
    type="date"
    value={date}
    onChange={(e)=>setDate(e.target.value)}
    className="border rounded-xl px-3 py-2"
  />

  <select
    value={category}
    onChange={(e)=>setCategory(e.target.value)}
    className="border rounded-xl px-3 py-2"
  >
    {expenseCategories.map(c=>(
      <option key={c}>{c}</option>
    ))}
  </select>

  <input
    type="text"
    placeholder="Description"
    value={description}
    onChange={(e)=>setDescription(e.target.value)}
    className="border rounded-xl px-3 py-2"
  />

  <input
    type="number"
    placeholder="Amount"
    value={amount}
    onChange={(e)=>setAmount(e.target.value)}
    className="border rounded-xl px-3 py-2"
  />

  <select
    value={paymentMode}
    onChange={(e)=>setPaymentMode(e.target.value)}
    className="border rounded-xl px-3 py-2"
  >
    <option>Cash</option>
    <option>UPI</option>
  </select>

  <button
     onClick={addExpense}
    className="bg-blue-600 text-white rounded-xl px-4 py-2 hover:bg-blue-700"
  >
    Add Expense
  </button>

  </div>

    <div className="mt-8 grid md:grid-cols-3 gap-4">

      <input
        type="text"
        placeholder="🔍 Search description..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        className="border border-gray-200 rounded-xl px-3 py-2"
      />

      <select
        value={filterCategory}
        onChange={(e) => setFilterCategory(e.target.value)}
        className="border border-gray-200 rounded-xl px-3 py-2"
      >
        <option>All</option>
        {expenseCategories.map((c) => (
          <option key={c}>{c}</option>
        ))}
      </select>

      <select
        value={filterPayment}
        onChange={(e) => setFilterPayment(e.target.value)}
        className="border border-gray-200 rounded-xl px-3 py-2"
      >
        <option>All</option>
        <option>Cash</option>
        <option>UPI</option>
      </select>

    </div>

    <div className="mt-8">

        <table className="w-full text-sm min-w-[700px]">

        <thead>
          <tr className="border-b border-gray-100">
            <th className="text-left py-3 px-3 text-xs font-semibold text-gray-500">Date</th>
            <th className="text-left py-3 px-3 text-xs font-semibold text-gray-500">Category</th>
            <th className="text-left py-3 px-3 text-xs font-semibold text-gray-500">Description</th>
            <th className="text-left py-3 px-3 text-xs font-semibold text-gray-500">Amount</th>
            <th className="text-left py-3 px-3 text-xs font-semibold text-gray-500">Payment</th>
            <th className="text-left py-3 px-3 text-xs font-semibold text-gray-500">Actions</th>
          </tr>
        </thead>

        <tbody>

        {filteredExpenses.map(exp=>(
        <tr key={exp.id} className="border-b border-gray-50 hover:bg-gray-50">

        <td className="p-3">{exp.date}</td>

        <td className="p-3">{exp.category}</td>

        <td className="p-3">{exp.description}</td>

        <td className="p-3 font-semibold text-red-600">
        ₹{(exp.amount ?? 0).toLocaleString("en-IN")}
        </td>

        <td className="p-3">
        {exp.paymentMode}
        </td>
    
        <td className="py-3 px-3">
          <div className="flex items-center gap-2">
          
          <button
            onClick={() => editExpense(exp)}
            className="px-3 py-1 rounded-lg bg-blue-50 text-blue-600 hover:bg-blue-100 transition"
          >
            ✏️ Edit
          </button>
          <button
            onClick={() => deleteExpense(exp.id)}
            className="text-red-600 hover:text-red-700 font-medium"
          >
            🗑 Delete
          </button>
       </div>
      </td>
      </tr>
    ))}
  </tbody>
</table>

</div>
</div>
</div>
  );
}

// ─── Reports ──────────────────────────────────────────────────────────────────

interface ReportsProps {
  earnings: Earning[]
  workers: Worker[]
  transactions: WTransaction[]
  expenses: Expense[]
}

function ReportsPage({ earnings, workers, transactions, expenses }: ReportsProps) {
  // FIX (Bug 1): "todayStr" was previously derived from the most recent
  // earning's date (earnings.sort().pop()), NOT the actual current date.
  // That meant the "Today" filter compared everything (earnings, expenses,
  // transactions) against a stale reference date whenever the latest earning
  // wasn't logged today - which is why expenses from a different day were
  // showing up under "Today". We now always use the real current date.
  const todayStr = new Date().toISOString().substring(0, 10);

  // ===== State =====
  const [reportFilter, setReportFilter] = useState("Today");
  const [fromDate, setFromDate] = useState(todayStr);
  const [toDate, setToDate] = useState(todayStr);
  const [filteredEarnings, setFilteredEarnings] = useState<Earning[]>([]);
  const [filteredExpenses, setFilteredExpenses] = useState<Expense[]>([]);
  const [customApplied, setCustomApplied] = useState(false);

  // ===== Shared period-filter helper =====
  // Both earnings and expenses need the exact same "which period is this row in"
  // logic, so it lives in one place instead of being duplicated (and drifting
  // out of sync, which is what caused the bugs before).
  function filterByPeriod<T extends { date: string }>(items: T[]): T[] {
    switch (reportFilter) {
      case "Today":
        return items.filter(i => i.date.substring(0, 10) === todayStr);

      case "This Week": {
        const start = new Date(todayStr);
        start.setHours(0, 0, 0, 0);

        // Monday as first day of week
        const day = start.getDay();
        const diff = day === 0 ? 6 : day - 1;
        start.setDate(start.getDate() - diff);

        const end = new Date(start);
        end.setDate(start.getDate() + 6);
        end.setHours(23, 59, 59, 999);

        return items.filter(i => {
          const d = new Date(i.date);
          return d >= start && d <= end;
        });
      }

      case "This Month":
        return items.filter(
          i => i.date.substring(0, 7) === todayStr.substring(0, 7)
        );

      case "This Year":
        return items.filter(
          i => i.date.substring(0, 4) === todayStr.substring(0, 4)
        );

      case "All Time":
        return items;

      case "Custom":
        return items.filter(i => {
          const d = i.date.substring(0, 10);
          return d >= fromDate && d <= toDate;
        });

      default:
        return items;
    }
  }

  function applyFilter() {
    setFilteredEarnings(filterByPeriod(earnings));
    setFilteredExpenses(filterByPeriod(expenses));
  }

  useEffect(() => {
    applyFilter();
    // Re-run whenever the underlying data changes, the period dropdown changes,
    // or the Custom From/To dates change and Generate is clicked (customApplied
    // toggles as a signal to re-filter even if from/to didn't change value-wise).
  }, [earnings, expenses, reportFilter, fromDate, toDate, customApplied]);

  // ===== Calculations =====
  const cashTotal = filteredEarnings
    .filter(e => e.paymentMode === "Cash")
    .reduce((s, e) => s + Number(e.amount), 0);

  const upiTotal = filteredEarnings
    .filter(e => e.paymentMode === "UPI")
    .reduce((s, e) => s + Number(e.amount), 0);

  const totalEarnings = cashTotal + upiTotal;

  const pendingPmt = activePending(filteredEarnings);

  const pendingTotal = pendingPmt.reduce(
    (s, e) => s + Number(e.amount),
    0
  );

  const pendingSal = workers.reduce(
    (s, w) => s + calcRemaining(w, transactions),
    0
  );

  const totalExpenses = filteredExpenses.reduce(
    (sum, e) => sum + (Number(e.amount) || 0),
    0
  );

  const totalSalary = workers.reduce(
    (sum, w) => sum + (Number(w.salary) || 0),
    0
  );

  const netProfit = totalEarnings - totalExpenses - totalSalary;

  const pieData = [
    { name: "Cash", value: cashTotal },
    { name: "UPI", value: upiTotal },
  ];

  const PIE_COLORS = ["#2563EB", "#F97316"];

  const monthlyMap: Record<string, number> = {};

  filteredEarnings.forEach((e) => {
    const month = new Date(e.date).toLocaleString(
      "en-IN",
      {
        month: "short",
        year: "numeric"
      }
    );

    monthlyMap[month] =
      (monthlyMap[month] || 0) + Number(e.amount);
  });

  const monthlyExpenseMap: Record<string, number> = {};

  filteredExpenses.forEach((e) => {
    const month = new Date(e.date).toLocaleString(
      "en-IN",
      {
        month: "short",
        year: "numeric"
      }
    );

    monthlyExpenseMap[month] =
      (monthlyExpenseMap[month] || 0) + Number(e.amount);
  });

  const monthlyData = Object.keys(monthlyMap).map(month => ({
    month,
    earnings: monthlyMap[month],
    expenses: monthlyExpenseMap[month] || 0,
  }));

  const dailyMap: Record<string, number> = {};

  filteredEarnings.forEach((e) => {
    dailyMap[e.date] = (dailyMap[e.date] || 0) + Number(e.amount);
  });

  const dailyData = Object.keys(dailyMap)
    .sort()
    .map((date) => ({
      day: date,
      earnings: dailyMap[date],
    }));

  const filteredTransactions = transactions.filter(t => {
    const d = t.date.substring(0, 10);

    switch (reportFilter) {
      case "Today":
        return d === todayStr;

      case "This Week": {
        const today = new Date(todayStr);

        const start = new Date(today);
        start.setHours(0, 0, 0, 0);

        // Monday as first day of week
        const day = start.getDay();
        const diff = day === 0 ? 6 : day - 1;

        start.setDate(start.getDate() - diff);

        const end = new Date(start);
        end.setDate(start.getDate() + 6);
        end.setHours(23, 59, 59, 999);

        return new Date(t.date) >= start && new Date(t.date) <= end;
      }

      case "This Month":
        return d.substring(0, 7) === todayStr.substring(0, 7);

      case "This Year":
        return d.substring(0, 4) === todayStr.substring(0, 4);

      case "Custom":
        return d >= fromDate && d <= toDate;

      default:
        return true;
    }
  });

  const deductionAmount = filteredTransactions
    .filter(t => t.type !== "Salary Settlement")
    .reduce((sum, t) => sum + Number(t.amount ?? 0), 0
    );

  const totalPaid = filteredTransactions.reduce(
    (sum, t) => sum + Number(t.amount ?? 0),
    0
  );

  function exportExcel() {
    const workbook = XLSX.utils.book_new();

    // Earnings Sheet
    const earningsSheet = XLSX.utils.json_to_sheet(
      filteredEarnings.map(e => ({
        Date: e.date,
        Customer: e.customerName,
        Vehicle: e.vehicleType,
        "Wash Type": e.washType,
        Amount: e.amount,
        Payment: e.paymentMode,
        Status: e.status,
        Remarks: e.remarks,
      }))
    );
    XLSX.utils.book_append_sheet(workbook, earningsSheet, "Earnings");

    // Expenses Sheet
    const expensesSheet = XLSX.utils.json_to_sheet(
      filteredExpenses.map(e => ({
        Date: e.date,
        Amount: e.amount,
        Category: (e as any).category ?? "",
        Notes: (e as any).notes ?? "",
      }))
    );
    XLSX.utils.book_append_sheet(workbook, expensesSheet, "Expenses");

    // Workers Sheet
    const workersSheet = XLSX.utils.json_to_sheet(
      workers.map(w => ({
        Name: w.name,
        Salary: w.salary,
        Status: w.status,
        JoinDate: w.joinDate,
      }))
    );
    XLSX.utils.book_append_sheet(workbook, workersSheet, "Workers");

    // Transactions Sheet
    const transactionSheet = XLSX.utils.json_to_sheet(transactions);
    XLSX.utils.book_append_sheet(workbook, transactionSheet, "Transactions");

    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    const file = new Blob([excelBuffer], {
      type:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    saveAs(file, "WashingStation_Report.xlsx");
  }

  function printReport() {
    try {
      const doc = new jsPDF();

      // ===== Report period label =====
      let reportPeriod = "";

      if (reportFilter === "Today") {
        reportPeriod = todayStr;
      } else if (reportFilter === "This Week") {
        reportPeriod = "This Week";
      } else if (reportFilter === "This Month") {
        reportPeriod = new Date(todayStr).toLocaleString("default", {
          month: "long",
          year: "numeric",
        });
      } else if (reportFilter === "This Year") {
        reportPeriod = todayStr.substring(0, 4);
      } else if (reportFilter === "All Time") {
        reportPeriod = "All Time";
      } else if (reportFilter === "Custom") {
        reportPeriod = `${fromDate} to ${toDate}`;
      }

      const settings = JSON.parse(localStorage.getItem("settings") || "{}");
      const businessName = settings.businessName || "SHREE WASH STATION";

      // ===== Header (drawn first, so it sits above the tables) =====
      doc.setFont("helvetica", "bold");
      doc.setFontSize(20);
      doc.text(businessName.toUpperCase(), 105, 18, { align: "center" });

      doc.setFontSize(14);
      doc.text("Business Report", 105, 28, { align: "center" });

      doc.setFont("helvetica", "normal");
      doc.setFontSize(10);
      doc.text(`Report Period: ${reportPeriod}`, 14, 38);
      doc.text(
        `Generated on: ${new Date().toLocaleDateString("en-IN")}`,
        14,
        44
      );

      // ===== Summary =====
      autoTable(doc, {
        startY: 50,
        theme: "grid",
        head: [["Metric", "Value"]],
        body: [
          ["Total Earnings", `Rs. ${totalEarnings.toLocaleString("en-IN")}`],
          ["Cash Collected", `Rs. ${cashTotal.toLocaleString("en-IN")}`],
          ["UPI Collected", `Rs. ${upiTotal.toLocaleString("en-IN")}`],
          ["Pending Payments", `Rs. ${pendingTotal.toLocaleString("en-IN")}`],
          ["Pending Salary", `Rs. ${pendingSal.toLocaleString("en-IN")}`],
          ["Worker Deductions", `Rs. ${deductionAmount.toLocaleString("en-IN")}`],
          ["Total Salary", `Rs. ${totalSalary.toLocaleString("en-IN")}`],
          ["Business Expenses", `Rs. ${totalExpenses.toLocaleString("en-IN")}`],
          ["Net Profit", `Rs. ${netProfit.toLocaleString("en-IN")}`],
        ],
        headStyles: {
          fillColor: [37, 99, 235],
        },
      });

      // ===== Earnings =====
      autoTable(doc, {
        startY: (doc as any).lastAutoTable.finalY + 10,
        theme: "striped",
        head: [[
          "Date",
          "Customer",
          "Vehicle",
          "Wash",
          "Amount",
          "Payment"
        ]],
        body: filteredEarnings.map(e => [
          e.date,
          e.customerName || "-",
          e.vehicleType,
          e.washType,
          "Rs. " + e.amount,
          e.paymentMode
        ]),
        headStyles: {
          fillColor: [22, 163, 74],
        },
      });

      // ===== Workers =====
      autoTable(doc, {
        startY: (doc as any).lastAutoTable.finalY + 10,
        theme: "striped",
        head: [["Worker", "Salary", "Status"]],
        body: workers.map(w => [
          w.name,
          "Rs. " + w.salary,
          w.status,
        ]),
        headStyles: {
          fillColor: [249, 115, 22],
        },
      });

      // ===== Footer =====
      doc.setFontSize(10);
      doc.text(
        "Generated by Washing Station Management System",
        105,
        285,
        { align: "center" }
      );

      doc.save("WashingStation_Report.pdf");

    } catch (err) {
      console.error(err);
      alert("Failed to generate PDF");
    }
  }

  return (
    <div className="p-4 lg:p-6 space-y-5">
      <div className="flex justify-between items-center flex-wrap gap-4">

        <div className="flex items-center gap-2">
          <label className="text-sm font-medium text-gray-600">
            Report Period:
          </label>

          <select
            value={reportFilter}
            onChange={(e) => setReportFilter(e.target.value)}
            className="border border-gray-300 rounded-lg px-3 py-2 text-sm"
          >
            <option>Today</option>
            <option>This Week</option>
            <option>This Month</option>
            <option>This Year</option>
            <option>All Time</option>
            <option>Custom</option>
          </select>
        </div>

        <div className="flex gap-3">
          <button
            onClick={exportExcel}
            className="border border-blue-600 text-blue-600 hover:bg-blue-50 text-sm font-semibold px-4 py-2.5 rounded-xl"
          >
            📥 Export Excel
          </button>
          <button
            onClick={printReport}
            className="border border-gray-200 text-gray-600 hover:bg-gray-50 text-sm font-semibold px-4 py-2.5 rounded-xl"
          >
            🖨️ Print Report
          </button>
        </div>
      </div>

      {reportFilter === "Custom" && (
        <div className="flex flex-wrap items-end gap-4 mt-4">

          <div>
            <label className="block text-sm font-medium text-gray-600 mb-1">
              From Date
            </label>
            <input
              type="date"
              value={fromDate}
              onChange={(e) => setFromDate(e.target.value)}
              className="border border-gray-300 rounded-lg px-3 py-2"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-600 mb-1">
              To Date
            </label>
            <input
              type="date"
              value={toDate}
              onChange={(e) => setToDate(e.target.value)}
              className="border border-gray-300 rounded-lg px-3 py-2"
            />
          </div>

          <div className="flex items-end gap-3">
            <button
              onClick={() => {
                setCustomApplied(prev => !prev);
              }}
              className="bg-blue-600 text-white px-5 py-2 rounded-lg"
            >
              🔍 Generate Report
            </button>

            <button
              onClick={() => {
                setReportFilter("Today");
                setFromDate(todayStr);
                setToDate(todayStr);
              }}
              className="border border-gray-300 hover:bg-gray-100 px-5 py-2 rounded-lg"
            >
              Reset
            </button>
          </div>
        </div>
      )}

      {/* Key metrics */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        {[
          { label: 'Total Earnings',     value: fmt(totalEarnings),        color: 'text-gray-900',   bg: 'bg-blue-50' },
          { label: 'Cash Collected',     value: fmt(cashTotal),            color: 'text-green-600',  bg: 'bg-green-50' },
          { label: 'UPI Collected',      value: fmt(upiTotal),             color: 'text-blue-600',   bg: 'bg-blue-50' },
          { label: 'Pending Payments',   value: fmt(pendingTotal),         color: 'text-orange-600', bg: 'bg-orange-50' },
          { label: 'Pending Salary',     value: fmt(pendingSal),           color: 'text-red-500',    bg: 'bg-red-50' },
          { label: 'Business Expenses',  value: fmt(totalExpenses),        color: 'text-red-600',    bg: 'bg-red-50' },
          { label: 'Net Profit',         value: fmt(netProfit),            color: netProfit >= 0 ? 'text-green-600' : 'text-red-600', bg: 'bg-purple-50' },
          { label: 'Active Workers',     value: String(workers.length),    color: 'text-gray-900',   bg: 'bg-purple-50' },
        ].map(r => (
          <div key={r.label} className={`${r.bg} rounded-xl p-4`}>
            <p className="text-xs font-medium text-gray-500 mb-1">{r.label}</p>
            <p className={`text-xl font-bold ${r.color}`}>{r.value}</p>
          </div>
        ))}
      </div>

      {/* Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <h3 className="font-semibold text-gray-900 mb-4">Daily Earnings</h3>
          <ResponsiveContainer width="100%" height={210}>
            <BarChart data={dailyData} margin={{ top: 4, right: 4, bottom: 0, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f3f4f6" />
              <XAxis dataKey="day" tick={{ fontSize: 11, fill: '#6b7280' }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: '#9ca3af' }} axisLine={false} tickLine={false} width={50} tickFormatter={v => `₹${(v/1000).toFixed(0)}k`} />
              <Tooltip formatter={(v) => [`₹${(Number(v)).toLocaleString('en-IN')}`, 'Earnings']} contentStyle={{ borderRadius: 12, border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.08)', fontSize: 13 }} />
              <Bar dataKey="earnings" fill="#2563EB" radius={[6, 6, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <h3 className="font-semibold text-gray-900 mb-4">Cash vs UPI</h3>
          <ResponsiveContainer width="100%" height={210}>
            <PieChart>
              <Pie data={pieData} cx="50%" cy="50%" innerRadius={65} outerRadius={90} paddingAngle={4} dataKey="value">
                {pieData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i]} />)}
              </Pie>
              <Tooltip formatter={(v) => `₹${(Number(v)).toLocaleString('en-IN')}`} contentStyle={{ borderRadius: 12, border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.08)', fontSize: 13 }} />
              <Legend iconType="circle" />
            </PieChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5 lg:col-span-2">
          <h3 className="font-semibold text-gray-900 mb-4">Monthly Reports</h3>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={monthlyData} margin={{ top: 4, right: 4, bottom: 0, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f3f4f6" />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: '#6b7280' }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: '#9ca3af' }} axisLine={false} tickLine={false} width={55} tickFormatter={v => `₹${(v/1000).toFixed(0)}k`} />
              <Tooltip formatter={(v) => `₹${(Number(v)).toLocaleString('en-IN')}`} contentStyle={{ borderRadius: 12, border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.08)', fontSize: 13 }} />
              <Legend iconType="circle" />
              <Bar dataKey="earnings" fill="#2563EB" radius={[4, 4, 0, 0]} name="Earnings" />
              <Bar dataKey="expenses" fill="#F97316" radius={[4, 4, 0, 0]} name="Expenses" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Pending payments summary */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
        <h3 className="font-semibold text-gray-900 mb-4">Pending Customer Payments</h3>
        {pendingPmt.length === 0 ? (
          <p className="text-gray-400 text-sm text-center py-6">No pending payments.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm min-w-[400px]">
              <thead>
                <tr className="border-b border-gray-100">
                  {['Customer', 'Vehicle', 'Amount', 'Date'].map(h => (
                    <th key={h} className="text-left py-2 px-3 text-xs font-semibold text-gray-500">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {pendingPmt.map(e => (
                  <tr key={e.id} className="border-b border-gray-50 hover:bg-gray-50">
                    <td className="py-2.5 px-3 font-medium text-gray-800">{e.customerName}</td>
                    <td className="py-2.5 px-3 text-gray-600">{e.vehicleType}</td>
                    <td className="py-2.5 px-3 font-bold text-orange-600">{fmt(e.amount)}</td>
                    <td className="py-2.5 px-3 text-gray-500 text-xs">{e.date}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr className="border-t border-gray-100">
                  <td colSpan={2} className="py-2.5 px-3 text-xs font-semibold text-gray-500">Total</td>
                  <td className="py-2.5 px-3 font-bold text-orange-600">{fmt(pendingTotal)}</td>
                  <td />
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </div>

      {/* Worker salary summary */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
        <h3 className="font-semibold text-gray-900 mb-4">Worker Salary Summary</h3>
        <div className="overflow-x-auto">
          <table className="w-full text-sm min-w-[500px]">
            <thead>
              <tr className="border-b border-gray-100">
                {['Worker', 'Monthly Salary', 'Deductions', 'Remaining', 'Status'].map(h => (
                  <th key={h} className="text-left py-2.5 px-3 text-xs font-semibold text-gray-500">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {workers.map(w => {
                const rem = calcRemaining(w, transactions)
                const dec = calcDeductions(w, transactions)
                return (
                  <tr key={w.id} className="border-b border-gray-50 hover:bg-gray-50">
                    <td className="py-3 px-3">
                      <div className="flex items-center gap-2.5">
                        <div className="w-7 h-7 rounded-full text-white text-xs font-bold flex items-center justify-center"
                          style={{ backgroundColor: w.color }}>{w.initials}</div>
                        <span className="font-medium text-gray-800">{w.name}</span>
                      </div>
                    </td>
                    <td className="py-3 px-3 text-gray-700">{fmt(w.salary)}</td>
                    <td className="py-3 px-3 text-red-500">{fmt(dec)}</td>
                    <td className="py-3 px-3 font-bold text-orange-600">{fmt(rem)}</td>
                    <td className="py-3 px-3">
                      <span className={`px-2.5 py-0.5 rounded-full text-xs font-semibold ${rem === 0 ? 'bg-green-100 text-green-700' : rem < w.salary * 0.3 ? 'bg-yellow-100 text-yellow-700' : 'bg-red-100 text-red-700'}`}>
                        {rem === 0 ? 'Fully Paid' : rem < w.salary * 0.3 ? 'Almost Done' : 'Pending'}
                      </span>
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

// ─── Settings ─────────────────────────────────────────────────────────────────
interface SettingsProps {
  workers: Worker[];
  earnings: Earning[];
  transactions: WTransaction[];
  settlements: SalarySettlement[];
}

function SettingsPage({
  workers,
  earnings,
  transactions,
  settlements,
}: SettingsProps) {
  const [form, setForm] = useState({
    businessName: 'Shree Wash Station', ownerName: 'Pavithra Kumar',
    phone: '9876543210', theme: 'Light', currentPassword: '', newPassword: '',
  })
  const [saved, setSaved] = useState(false)
  const fc = 'w-full border border-gray-200 rounded-xl px-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white'

  useEffect(() => {
  const saved = localStorage.getItem("settings");

  if (saved) {
    setForm(JSON.parse(saved));
  }
  }, []);

  function exportBackup() {
  const backup = {
    settings: form,
    workers,
    earnings,
    transactions,
    settlements,
    exportedAt: new Date().toISOString(),
  };

  const blob = new Blob(
    [JSON.stringify(backup, null, 2)],
    {
      type: "application/json",
    }
  );

  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");

  a.href = url;

  a.download = `washing_station_backup_${
    new Date().toISOString().split("T")[0]
  }.json`;

  a.click();

  URL.revokeObjectURL(url);
  }

  return (
    <div className="p-4 lg:p-6">
      <div className="max-w-2xl mx-auto space-y-5">
        {saved && <div className="bg-green-50 border border-green-200 text-green-700 rounded-xl px-4 py-3 text-sm font-medium">✅ Settings saved!</div>}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <h3 className="font-semibold text-gray-900 mb-4">Business Information</h3>
          <div className="space-y-4">
            {[{ label: 'Business Name', key: 'businessName', placeholder: 'Shree Wash Station' },
              { label: 'Owner Name', key: 'ownerName', placeholder: 'Pavithra Kumar' },
              { label: 'Phone Number', key: 'phone', placeholder: '9876543210' }].map(f => (
              <div key={f.key}>
                <label className="text-xs font-medium text-gray-600 mb-1.5 block">{f.label}</label>
                <input type="text" value={form[f.key as keyof typeof form]} onChange={e => setForm(p => ({ ...p, [f.key]: e.target.value }))}
                  placeholder={f.placeholder} className={fc} />
              </div>
            ))}
          </div>
        </div>
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-5">
          <h3 className="font-semibold text-gray-900 mb-4">Change Password</h3>
          <div className="space-y-4">
            <div>
              <label className="text-xs font-medium text-gray-600 mb-1.5 block">Current Password</label>
              <input type="password" value={form.currentPassword} onChange={e => setForm(f => ({ ...f, currentPassword: e.target.value }))} placeholder="••••••••" className={fc} />
            </div>
            <div>
              <label className="text-xs font-medium text-gray-600 mb-1.5 block">New Password</label>
              <input type="password" value={form.newPassword} onChange={e => setForm(f => ({ ...f, newPassword: e.target.value }))} placeholder="••••••••" className={fc} />
            </div>
          </div>
        </div>
                <button
          onClick={async () => {
          try {
            const response = await fetch(
              "https://washing-station-production.up.railway.app/api/auth/change-password",
              {
                method: "PUT",
                headers: {
                  "Content-Type": "application/json",
                },
                body: JSON.stringify({
                  username: "admin",
                  currentPassword: form.currentPassword,
                  newPassword: form.newPassword,
                }),
              }
            );
            const result = await response.text();
            if (result === "SUCCESS") {
              alert("Password changed successfully!");
            } else if (result === "INVALID_PASSWORD") {
              alert("Current password is incorrect.");
              return;
            }
            localStorage.setItem("settings", JSON.stringify(form));
            setSaved(true);
            setTimeout(() => setSaved(false), 3000);
          } catch (err) {
            console.error(err);
            alert("Unable to connect to server.");
          }
                  }}
          className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3.5 rounded-xl transition-colors shadow-sm"
        >
          Save Settings
        </button>
      </div>
    </div>
  );
}
  


// ─── Login ────────────────────────────────────────────────────────────────────

function LoginPage({ onLogin }: { onLogin: () => void }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [remember, setRemember] = useState(false)
  const [showPassword, setShowPassword] = useState(false)
  async function handleLogin() {
    try {
      const response = await fetch("https://washing-station-production.up.railway.app/api/auth/login", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          username,
          password,
        }),
      });

      const result = await response.text();

      if (result === "SUCCESS") {

        if (remember) {
          localStorage.setItem("loggedUser", username);
          localStorage.setItem("isLoggedIn", "true");
        } else {
          sessionStorage.setItem("loggedUser", username);
          sessionStorage.setItem("isLoggedIn", "true");
        }
        onLogin();
      } else {
        alert("Invalid username or password");
      }
    } catch (err) {
      console.error(err);
      alert("Unable to connect to server");
    }
  }
  const fc = 'w-full border border-gray-200 rounded-xl px-4 py-3 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white transition-shadow'

  return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center p-6">
      <div className="w-full max-w-md bg-white rounded-2xl shadow-sm border border-gray-100 p-8">
        {/* Header */}
        <div className="mb-7">
          <div className="flex items-center gap-3 mb-3">
            <div className="w-11 h-11 bg-blue-600 rounded-xl flex items-center justify-center text-white font-bold text-xl flex-shrink-0">W</div>
            <div>
              <h1 className="text-xl font-bold text-gray-900">Wash Station</h1>
              <p className="text-xs text-blue-600 font-semibold tracking-wide uppercase">Management System</p>
            </div>
          </div>
          <p className="text-gray-500 text-sm leading-relaxed">Manage workers, daily earnings and expenses efficiently.</p>
        </div>

        {/* Form */}
        <div className="space-y-4">
          <div>
            <label className="text-xs font-medium text-gray-600 mb-1.5 block">
              Username
            </label>

            <input
              type="text"
              value={username}
              onChange={e => setUsername(e.target.value)}
              placeholder="admin"
              onKeyDown={e => e.key === "Enter" && handleLogin()}
              className={fc}
            />
          </div>
          </div>
          <div>
          <label className="text-xs font-medium text-gray-600 mb-1.5 block">
            Password
          </label>

          <div className="relative">
            <input
              type={showPassword ? "text" : "password"}
              value={password}
              onChange={e => setPassword(e.target.value)}
              placeholder="••••••••"
              onKeyDown={e => e.key === "Enter" && handleLogin()}
              className={`${fc} pr-12`}
            />

            <button
              type="button"
              onClick={() => setShowPassword(!showPassword)}
              className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-gray-700"
            >
              {showPassword ? "🙈" : "👁️"}
            </button>
          </div>
        </div>

        {/* Dedication */}
        <div className="mt-7 pt-5 border-t border-blue-100">
          <p className="text-[11px] text-gray-400 text-center leading-relaxed">
            This Wash Station Management System has been specially designed and developed
            exclusively for{' '}
            <span className="font-semibold text-blue-600">Mr. Pavithra Kumar</span>
            {' '}to simplify daily operations, worker management, salary settlements,
            daily earnings tracking, pending customer payments, and business reporting.
          </p>
        </div>
      </div>
    </div>
  )
}

// ─── App ──────────────────────────────────────────────────────────────────────

const PAGE_TITLES: Record<string, string> = {
  dashboard:          'Dashboard',
  earnings:           'Daily Earnings',
  'pending-payments': 'Pending Payments',
  workers:            'Workers',
  'worker-detail':    'Worker Details',
  'add-transaction':  'Worker Transactions',
  'salary-settlement':'Salary Settlement',
  expenses:           'Expenses',
  reports:            'Reports',
  settings:           'Settings',
}

export default function App() {

  const [page, setPage] = useState<Page>(() => {
  const logged =
    localStorage.getItem("isLoggedIn") ||
    sessionStorage.getItem("isLoggedIn");

  return logged ? "dashboard" : "login";
});
  const [collapsed,  setCollapsed]  = useState(false)
  const [mobileOpen, setMobileOpen] = useState(false)
  const [selWorker,  setSelWorker]  = useState<Worker | null>(null)
  const [expenses, setExpenses] = useState<Expense[]>([]);

  // ── Single source of truth ──────────────────────────────────────────────────
  const [earnings, setEarnings] = useState<Earning[]>([])
  const [workers, setWorkers] = useState<Worker[]>([])
  const [transactions, setTransactions] = useState<WTransaction[]>([])
  useEffect(() => {
  fetch("https://washing-station-production.up.railway.app/api/workers")
    .then(res => res.json())
    .then(data => {
      const COLORS = [
        "#2563EB",
        "#F97316",
        "#10B981",
        "#8B5CF6",
        "#EC4899",
        "#14B8A6"
      ];

      const workers = data.map((worker: any, index: number) => ({
        ...worker,
        initials: worker.name
          .split(" ")
          .map((n: string) => n[0])
          .join("")
          .toUpperCase()
          .slice(0, 2),
        color: COLORS[index % COLORS.length],
      }));

      setWorkers(workers);
    })
    .catch(err => console.error(err));
}, []);
useEffect(() => {
  fetch("https://washing-station-production.up.railway.app/api/earnings")
    .then(res => res.json())
    .then(data => {
      setEarnings(data);
    })
    .catch(err => console.error(err));
}, []);
useEffect(() => {
  fetch("https://washing-station-production.up.railway.app/api/dashboard")
    .then((res) => res.json())
    .then((data) => {
      setDashboard(data);
    })
    .catch((err) => console.error(err));
}, []);

useEffect(() => {
  fetch("https://washing-station-production.up.railway.app/api/expenses")
    .then(res => res.json())
    .then(data => setExpenses(data))
    .catch(err => console.error(err));
}, []);

const loadTransactions = () => {
  fetch("https://washing-station-production.up.railway.app/api/transactions")
    .then(res => res.json())
    .then(data => {
      const txs = data.map((tx: any) => ({
        id: tx.id,
        workerId: tx.worker.id,
        workerName: tx.worker.name,
        amount: tx.amount,
        date: tx.date,
        type: tx.type,
        remarks: tx.remarks,
      }));

      setTransactions(txs);
    })
    .catch(err => console.error(err));
};

useEffect(() => {
  loadTransactions();
}, []);

useEffect(() => {
  fetch("https://washing-station-production.up.railway.app/api/salary-settlements")
    .then(res => res.json())
    .then(data => {
      const mapped = data.map((s: any) => ({
        workerId: s.worker.id,
        month: s.month,
        year: s.year,
        settled: true,
        settlementDate: s.settlementDate,
      }));

      setSettlements(mapped);
    })
    .catch(err => console.error(err));
}, []);
 const [settlements, setSettlements] = useState<SettlementRecord[]>([])
  const [dashboard, setDashboard] = useState({
  totalWorkers: 0,
  activeWorkers: 0,
  totalSalary: 0,
  totalTransactions: 0,
});

const totalEarnings = earnings.reduce(
  (sum, e) => sum + (e.amount ?? 0),
  0
);

const totalExpenses = expenses.reduce(
  (sum, e) => sum + (e.amount ?? 0),
  0
);

const netProfit = totalEarnings - totalExpenses;

  function handleSelectWorker(w: Worker) { setSelWorker(w); setPage('worker-detail') }

  if (page === 'login') return <LoginPage onLogin={() => setPage('dashboard')} />

  return (
    <div className="flex h-screen bg-gray-50 overflow-hidden">
      <Sidebar
        current={page} onNav={setPage}
        collapsed={collapsed} onToggle={() => setCollapsed(v => !v)}
        mobileOpen={mobileOpen} onMobileClose={() => setMobileOpen(false)}
        onLogout={() => {
          localStorage.removeItem("loggedUser");
          localStorage.removeItem("isLoggedIn");

          sessionStorage.removeItem("loggedUser");
          sessionStorage.removeItem("isLoggedIn");

          setPage("login");
        }}
      />
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
        <TopBar title={PAGE_TITLES[page] || 'Wash Station'} onMenuClick={() => setMobileOpen(v => !v)} />
        <main className="flex-1 overflow-y-auto">
          {page === 'dashboard'          && <DashboardPage earnings={earnings} workers={workers} transactions={transactions} dashboard={dashboard} onNav={setPage} onSelectWorker={handleSelectWorker} />}
          {page === 'earnings'           && <DailyEarningsPage earnings={earnings} setEarnings={setEarnings} />}
          {page === 'pending-payments'   && <PendingPaymentsPage earnings={earnings} setEarnings={setEarnings} />}
          {page === 'workers'            && <WorkersPage workers={workers} setWorkers={setWorkers} transactions={transactions} onSelectWorker={handleSelectWorker} />}
          {page === 'worker-detail'      && selWorker && <WorkerDetailPage worker={selWorker} transactions={transactions} onBack={() => setPage('workers')} />}
          {page === 'add-transaction'    && <AddTransactionPage workers={workers} transactions={transactions} setTransactions={setTransactions} />}
          {page === 'salary-settlement'  && <SalarySettlementPage workers={workers} transactions={transactions} settlements={settlements} setSettlements={setSettlements} loadTransactions={loadTransactions}/>}
          {page === 'expenses'           && <ExpensesPage expenses={expenses} setExpenses={setExpenses}/>}
          {page === 'reports'            && <ReportsPage earnings={earnings} workers={workers} transactions={transactions} expenses={expenses}/>}
          {page === 'settings'           && <SettingsPage workers={workers} earnings={earnings} transactions={transactions} settlements={settlements}/>}
        </main>
      </div>
    </div>
  )
}