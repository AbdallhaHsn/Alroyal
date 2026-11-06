import React, { useState, useMemo } from 'react';
import {
  BookOpen,
  Calendar,
  FileSpreadsheet,
  Printer,
  ChevronDown,
  ChevronLeft,
  Search,
  Filter,
  CheckSquare,
  Square,
  X,
  Coins,
  TrendingUp,
  TrendingDown,
  Layers,
  ArrowDownLeft,
  CheckCircle2,
  RefreshCw,
  FolderTree,
  ExternalLink,
  Scale
} from 'lucide-react';

export interface GeneralAccountItem {
  code: string;
  name: string;
  parentCode: string;
  parentName: string;
  category: string;
}

export interface GeneralLedgerTransaction {
  Id: number;
  ReqNum: string;
  DocType: string;
  TheDate: string;
  AccountNumber: string;
  AccountName: string;
  Debtor: number;
  Creditor: number;
  DebtorDolar: number;
  CreditorDolar: number;
  Detail: string;
  Curruncy_Name: string;
  Check_Number: string;
  runningNet: number;
  runningNetDolar: number;
}

export interface GeneralLedgerChildAccount {
  AccountNumber: string;
  AccountName: string;
  openDebtor: number;
  openCreditor: number;
  openDebtorDolar: number;
  openCreditorDolar: number;
  periodDebtor: number;
  periodCreditor: number;
  periodDebtorDolar: number;
  periodCreditorDolar: number;
  totalDebtor: number;
  totalCreditor: number;
  totalDebtorDolar: number;
  totalCreditorDolar: number;
  netOpen: number;
  netOpenDolar: number;
  netEnding: number;
  netEndingDolar: number;
}

export interface GeneralLedgerSubAccountData {
  accountNumber: string;
  accountName: string;
  openDebtor: number;
  openCreditor: number;
  openDebtorDolar: number;
  openCreditorDolar: number;
  netOpen: number;
  netOpenDolar: number;
  periodDebtor: number;
  periodCreditor: number;
  periodDebtorDolar: number;
  periodCreditorDolar: number;
  totalDebtor: number;
  totalCreditor: number;
  netEnding: number;
  netEndingDolar: number;
  transactions: GeneralLedgerTransaction[];
}

export interface GeneralLedgerAccountData {
  code: string;
  name: string;
  parentName: string;
  category: string;
  openDebtor: number;
  openCreditor: number;
  openDebtorDolar: number;
  openCreditorDolar: number;
  netOpen: number;
  netOpenDolar: number;
  periodDebtor: number;
  periodCreditor: number;
  periodDebtorDolar: number;
  periodCreditorDolar: number;
  totalDebtor: number;
  totalCreditor: number;
  totalDebtorDolar: number;
  totalCreditorDolar: number;
  netEnding: number;
  netEndingDolar: number;
  transactionsCount?: number;
  transactions?: GeneralLedgerTransaction[];
  subAccountsCount?: number;
  subAccounts?: GeneralLedgerSubAccountData[];
  childAccountsCount?: number;
  childAccounts?: GeneralLedgerChildAccount[];
}

export interface GeneralLedgerResponse {
  startDate: string;
  endDate: string;
  mode: 'detailed' | 'summary';
  accountsCount: number;
  grandTotals: {
    openDebtor: number;
    openCreditor: number;
    openDebtorDolar: number;
    openCreditorDolar: number;
    netOpenDebit: number;
    netOpenCredit: number;
    netOpenDebitDolar: number;
    netOpenCreditDolar: number;
    periodDebtor: number;
    periodCreditor: number;
    periodDebtorDolar: number;
    periodCreditorDolar: number;
    totalDebtor: number;
    totalCreditor: number;
    totalDebtorDolar: number;
    totalCreditorDolar: number;
    netEndingDebit: number;
    netEndingCredit: number;
    netEndingDebitDolar: number;
    netEndingCreditDolar: number;
  };
  accounts: GeneralLedgerAccountData[];
}

interface GeneralLedgerViewProps {
  startDate: string;
  setStartDate: (d: string) => void;
  endDate: string;
  setEndDate: (d: string) => void;
  currency: 'EGP' | 'USD' | 'BOTH';
  setCurrency: (c: 'EGP' | 'USD' | 'BOTH') => void;
  mode: 'detailed' | 'summary';
  setMode: (m: 'detailed' | 'summary') => void;
  generalAccounts: GeneralAccountItem[];
  selectedCodes: string[];
  toggleAccount: (code: string) => void;
  selectAllAccounts: () => void;
  clearAllAccounts: () => void;
  data: GeneralLedgerResponse | null;
  loading: boolean;
  onRefresh: () => void;
  onExportExcel: () => void;
  formatCurrency: (val: number | undefined) => string;
  renderNetBadge: (val: number, isEnding?: boolean, symbol?: string) => React.ReactNode;
  openVoucherModal: (reqNum: string) => void;
}

export default function GeneralLedgerView({
  startDate,
  setStartDate,
  endDate,
  setEndDate,
  currency,
  setCurrency,
  mode,
  setMode,
  generalAccounts,
  selectedCodes,
  toggleAccount,
  selectAllAccounts,
  clearAllAccounts,
  data,
  loading,
  onRefresh,
  onExportExcel,
  formatCurrency,
  renderNetBadge,
  openVoucherModal
}: GeneralLedgerViewProps) {
  const [pickerOpen, setPickerOpen] = useState(false);
  const [searchAccount, setSearchAccount] = useState('');
  const [categoryFilter, setCategoryFilter] = useState<string>('all');
  
  // Track expanded cards and sub-accounts. Default is COLLAPSED!
  const [expandedCards, setExpandedCards] = useState<{ [code: string]: boolean }>({});
  const [expandedSubAccounts, setExpandedSubAccounts] = useState<{ [subCode: string]: boolean }>({});

  const toggleCollapse = (code: string) => {
    setExpandedCards(prev => ({ ...prev, [code]: !prev[code] }));
  };

  const toggleSubAccount = (subCode: string) => {
    setExpandedSubAccounts(prev => ({ ...prev, [subCode]: !prev[subCode] }));
  };

  const expandAll = () => {
    const all: { [k: string]: boolean } = {};
    (data?.accounts || []).forEach(a => { all[a.code] = true; });
    setExpandedCards(all);

    const allSub: { [k: string]: boolean } = {};
    (data?.accounts || []).forEach(a => {
      (a.subAccounts || []).forEach(s => {
        allSub[s.accountNumber] = true;
      });
    });
    setExpandedSubAccounts(allSub);
  };

  const collapseAll = () => {
    setExpandedCards({});
    setExpandedSubAccounts({});
  };

  const expandAllSubAccounts = (acc: GeneralLedgerAccountData) => {
    setExpandedSubAccounts(prev => {
      const next = { ...prev };
      (acc.subAccounts || []).forEach(s => {
        next[s.accountNumber] = true;
      });
      return next;
    });
  };

  const collapseAllSubAccounts = (acc: GeneralLedgerAccountData) => {
    setExpandedSubAccounts(prev => {
      const next = { ...prev };
      (acc.subAccounts || []).forEach(s => {
        delete next[s.accountNumber];
      });
      return next;
    });
  };

  // Filter general accounts inside the selector popover
  const filteredAccounts = useMemo(() => {
    let list = generalAccounts;
    if (categoryFilter !== 'all') {
      list = list.filter(a => a.category === categoryFilter || a.code.startsWith(categoryFilter));
    }
    if (searchAccount.trim()) {
      const q = searchAccount.toLowerCase().trim();
      list = list.filter(a => 
        a.name.toLowerCase().includes(q) || 
        a.code.includes(q) || 
        (a.parentName && a.parentName.toLowerCase().includes(q))
      );
    }
    return list;
  }, [generalAccounts, categoryFilter, searchAccount]);

  const formatDate = (dateStr: string) => {
    if (!dateStr) return '-';
    try {
      return new Date(dateStr).toISOString().split('T')[0];
    } catch {
      return dateStr;
    }
  };

  const gt = data?.grandTotals;

  return (
    <div className="space-y-4">
      {/* 1. FILTER & CONFIGURATION CONTROL PANEL */}
      <div className="enterprise-card rounded-2xl p-6 no-print">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4 mb-4">
          <div>
            <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <BookOpen className="w-5 h-5 text-blue-600" />
              <span>تقرير كشف الحسابات العامة (General Ledger Statement)</span>
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              كشف حركة تفصيلية وتجميعية للحسابات العامة المحددة، مع الأرصدة الافتتاحية وحركات القيود والأرصدة التراكمية المتحركة.
            </p>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={onExportExcel}
              disabled={loading || !data?.accounts?.length}
              className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white px-4 py-2 rounded-xl text-xs font-bold transition shadow-xs cursor-pointer"
            >
              <FileSpreadsheet className="w-4 h-4" />
              <span>تصدير Excel منسق (XLSX)</span>
            </button>
            <button
              onClick={() => window.print()}
              disabled={loading || !data?.accounts?.length}
              className="flex items-center gap-1.5 bg-slate-900 hover:bg-slate-800 disabled:opacity-50 text-white px-4 py-2 rounded-xl text-xs font-bold transition shadow-xs cursor-pointer"
            >
              <Printer className="w-4 h-4" />
              <span>طباعة كشف الحساب</span>
            </button>
          </div>
        </div>

        {/* Toolbar Controls Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3 pt-1">
          {/* Start Date */}
          <div>
            <label className="text-[11px] font-bold text-slate-600 block mb-1.5">
              من تاريخ الفترة:
            </label>
            <div className="relative">
              <Calendar className="w-4 h-4 absolute right-3 top-2.5 text-slate-400" />
              <input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="w-full bg-slate-50 border border-slate-300 rounded-xl pr-9 pl-3 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-blue-600 font-mono font-bold"
              />
            </div>
          </div>

          {/* End Date */}
          <div>
            <label className="text-[11px] font-bold text-slate-600 block mb-1.5">
              إلى تاريخ الفترة:
            </label>
            <div className="relative">
              <Calendar className="w-4 h-4 absolute right-3 top-2.5 text-slate-400" />
              <input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="w-full bg-slate-50 border border-slate-300 rounded-xl pr-9 pl-3 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-blue-600 font-mono font-bold"
              />
            </div>
          </div>

          {/* Currency Toggle */}
          <div>
            <label className="text-[11px] font-bold text-slate-600 block mb-1.5">
              عملة التقرير والعرض:
            </label>
            <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200">
              <button
                type="button"
                onClick={() => setCurrency('EGP')}
                className={`flex-1 py-1 text-xs rounded-lg font-bold transition ${
                  currency === 'EGP' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                جنيه سوداني (SDG)
              </button>
              <button
                type="button"
                onClick={() => setCurrency('USD')}
                className={`flex-1 py-1 text-xs rounded-lg font-bold transition ${
                  currency === 'USD' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                دولار ($)
              </button>
              <button
                type="button"
                onClick={() => setCurrency('BOTH')}
                className={`flex-1 py-1 text-xs rounded-lg font-bold transition ${
                  currency === 'BOTH' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                كلاهما
              </button>
            </div>
          </div>

          {/* Mode Toggle: Detailed vs Summary */}
          <div>
            <label className="text-[11px] font-bold text-slate-600 block mb-1.5">
              مستوى التفاصيل والعرض:
            </label>
            <div className="flex bg-slate-100 p-1 rounded-xl border border-slate-200">
              <button
                type="button"
                onClick={() => setMode('detailed')}
                className={`flex-1 py-1 text-xs rounded-lg font-bold transition ${
                  mode === 'detailed' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                تفصيلي (بالقيود)
              </button>
              <button
                type="button"
                onClick={() => setMode('summary')}
                className={`flex-1 py-1 text-xs rounded-lg font-bold transition ${
                  mode === 'summary' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                تجميعي (حسابات)
              </button>
            </div>
          </div>
        </div>

        {/* Account Multi-Selection Bar */}
        <div className="mt-4 pt-3 border-t border-slate-100">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-2">
            <div className="flex items-center gap-2">
              <label className="text-[11px] font-bold text-slate-700">
                الحسابات العامة المستهدفة:
              </label>
              <span className="font-mono bg-blue-100 text-blue-800 text-[11px] px-2 py-0.5 rounded-full font-bold">
                {selectedCodes.length} حساب مختار
              </span>
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => setPickerOpen(!pickerOpen)}
                className="flex items-center gap-1.5 bg-blue-50 text-blue-700 hover:bg-blue-100 px-3 py-1.5 rounded-lg text-xs font-bold transition border border-blue-200 shadow-2xs cursor-pointer"
              >
                <Filter className="w-3.5 h-3.5" />
                <span>{pickerOpen ? 'إغلاق قائمة الحسابات' : 'تحديد واختيار الحسابات العامة'}</span>
                <ChevronDown className={`w-3.5 h-3.5 transition-transform ${pickerOpen ? 'rotate-180' : ''}`} />
              </button>

              {selectedCodes.length > 0 && (
                <button
                  type="button"
                  onClick={clearAllAccounts}
                  className="text-xs text-rose-600 hover:text-rose-800 font-bold px-2 py-1 transition cursor-pointer"
                >
                  إلغاء التحديد
                </button>
              )}
            </div>
          </div>

          {/* Selected Account Tags */}
          <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto p-1 bg-slate-50/70 rounded-xl border border-slate-200/80">
            {selectedCodes.length === 0 ? (
              <span className="text-xs text-slate-400 p-1">
                لم يتم اختيار أي حساب عام. انقر على زر "تحديد واختيار الحسابات العامة" أعلاه لاختيار الحسابات.
              </span>
            ) : (
              selectedCodes.map(code => {
                const acc = generalAccounts.find(a => a.code === code);
                const name = acc?.name || code;
                return (
                  <span
                    key={code}
                    className="inline-flex items-center gap-1.5 bg-white border border-slate-300 text-slate-800 px-2.5 py-1 rounded-lg text-xs font-sans font-semibold shadow-2xs group"
                  >
                    <span className="font-mono text-blue-600 font-bold">{code}</span>
                    <span className="text-slate-800">{name}</span>
                    <button
                      type="button"
                      onClick={() => toggleAccount(code)}
                      className="text-slate-400 hover:text-rose-600 rounded-full p-0.5 transition"
                      title="إزالة هذا الحساب من التقرير"
                    >
                      <X className="w-3 h-3" />
                    </button>
                  </span>
                );
              })
            )}
          </div>

          {/* Popover / Collapsible Accounts Picker */}
          {pickerOpen && (
            <div className="mt-3 p-4 bg-white rounded-xl border border-blue-200 shadow-lg space-y-3 animate-in fade-in slide-in-from-top-2 duration-150">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-100">
                {/* Search */}
                <div className="relative flex-1">
                  <Search className="w-4 h-4 absolute right-3 top-2.5 text-slate-400" />
                  <input
                    type="text"
                    placeholder="ابحث عن كود الحساب أو اسمه..."
                    value={searchAccount}
                    onChange={(e) => setSearchAccount(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg pr-9 pl-3 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-blue-500"
                  />
                </div>

                {/* Category Pills */}
                <div className="flex items-center gap-1 overflow-x-auto text-xs">
                  {['all', 'الأصول', 'الخصوم', 'المصروفات', 'الإيرادات'].map(cat => (
                    <button
                      key={cat}
                      type="button"
                      onClick={() => setCategoryFilter(cat)}
                      className={`px-2.5 py-1 rounded-lg font-semibold whitespace-nowrap transition cursor-pointer ${
                        categoryFilter === cat 
                          ? 'bg-blue-600 text-white' 
                          : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      {cat === 'all' ? 'كافة التصنيفات' : cat}
                    </button>
                  ))}
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={selectAllAccounts}
                    className="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-3 py-1.5 rounded-lg transition cursor-pointer whitespace-nowrap"
                  >
                    تحديد كل المفلتر ({filteredAccounts.length})
                  </button>
                </div>
              </div>

              {/* Accounts Checkbox Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2 max-h-64 overflow-y-auto p-1">
                {filteredAccounts.map(a => {
                  const isChecked = selectedCodes.includes(a.code);
                  return (
                    <div
                      key={a.code}
                      onClick={() => toggleAccount(a.code)}
                      className={`flex items-center gap-2.5 p-2 rounded-lg border text-xs cursor-pointer select-none transition ${
                        isChecked 
                          ? 'bg-blue-50/70 border-blue-400 text-blue-900 shadow-2xs font-bold' 
                          : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
                      }`}
                    >
                      {isChecked ? (
                        <CheckSquare className="w-4 h-4 text-blue-600 flex-shrink-0" />
                      ) : (
                        <Square className="w-4 h-4 text-slate-400 flex-shrink-0" />
                      )}
                      <span className="font-mono text-blue-700 font-bold flex-shrink-0">
                        {a.code}
                      </span>
                      <span className="truncate flex-1 font-sans">
                        {a.name}
                      </span>
                      <span className="text-[10px] bg-slate-100 text-slate-500 px-1.5 py-0.2 rounded font-sans font-normal flex-shrink-0">
                        {a.category}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* 3. PRINTABLE OFFICIAL LETTERHEAD */}
      <div className="hidden print:block pb-2 mb-3 border-b-2 border-slate-800 text-slate-900 w-full max-w-full box-border">
        <div className="flex items-center justify-between gap-2 w-full">
          <div className="text-right shrink-0">
            <h1 className="text-base font-black tracking-tight text-slate-900">الرويال للتطوير العقاري</h1>
            <p className="text-[10px] font-bold text-slate-700">Al Royal Real Estate Development</p>
            <p className="text-[9px] text-slate-500 font-sans">الإدارة المالية المركزية • ERP</p>
          </div>

          <div className="text-center px-1">
            <div className="inline-block bg-slate-100 border border-slate-300 px-3 py-1 rounded">
              <h2 className="text-xs font-black text-slate-900">
                تقرير كشف الحسابات العامة (General Ledger)
              </h2>
              <span className="text-[10px] font-bold text-blue-800 font-sans block mt-0.5">
                {mode === 'detailed' ? 'كشف تفصيلي بالقيود والحركات التحليلية' : 'كشف تجميعي بالحسابات الفرعية'}
              </span>
            </div>
          </div>

          <div className="text-left font-mono text-[9px] bg-slate-50 border border-slate-300 p-1.5 rounded leading-tight shrink-0">
            <div><span className="text-slate-500 font-sans">الفترة:</span> <strong>{startDate}</strong> إلى <strong>{endDate}</strong></div>
            <div><span className="text-slate-500 font-sans">العملة:</span> <strong>{currency === 'EGP' ? 'جنيه سوداني (ج.س)' : currency === 'USD' ? 'دولار أمريكي ($)' : 'العملتان معاً'}</strong></div>
            <div><span className="text-slate-500 font-sans">تاريخ الطباعة:</span> <strong>{new Date().toLocaleDateString('ar-EG')}</strong></div>
          </div>
        </div>
      </div>

      {/* 4. EXPAND / COLLAPSE ALL TOOLBAR */}
      {data && data.accounts.length > 0 && (
        <div className="flex items-center justify-between text-xs px-2 no-print">
          <div className="text-slate-500 font-bold">
            نتائج كشف الحساب: <span className="text-blue-700 font-mono">{data.accounts.length}</span> حساب عام
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={expandAll}
              className="text-blue-600 hover:text-blue-800 font-bold cursor-pointer"
            >
              توسيع كافة الحسابات
            </button>
            <span className="text-slate-300">|</span>
            <button
              onClick={collapseAll}
              className="text-slate-500 hover:text-slate-700 font-bold cursor-pointer"
            >
              طي كافة الحسابات
            </button>
          </div>
        </div>
      )}

      {/* 5. ACCOUNTS STATEMENTS LIST */}
      {loading ? (
        <div className="enterprise-card rounded-2xl p-16 text-center text-slate-400">
          <RefreshCw className="w-8 h-8 animate-spin mx-auto text-blue-600 mb-3" />
          <p className="font-semibold text-slate-700 text-sm">جاري تجميع حركات وأرصدة كشف الحسابات العامة من قاعدة البيانات...</p>
        </div>
      ) : !data || data.accounts.length === 0 ? (
        <div className="enterprise-card rounded-2xl p-16 text-center text-slate-400">
          <BookOpen className="w-12 h-12 mx-auto text-slate-300 mb-3 stroke-[1.2]" />
          <h3 className="font-bold text-slate-700 text-sm mb-1">لا توجد بيانات مطابقة للعرض</h3>
          <p className="text-xs text-slate-500 max-w-md mx-auto">
            يرجى تحديد حساب عام واحد على الأقل من القائمة بالأعلى لعرض كشف الحساب الخاص به وتفاصيل القيود والحركات.
          </p>
        </div>
      ) : (
        <div className="space-y-6">
          {data.accounts.map((acc) => {
            const isCollapsed = !expandedCards[acc.code];
            const netOpenVal = currency === 'USD' ? acc.netOpenDolar : acc.netOpen;
            const netEndingVal = currency === 'USD' ? acc.netEndingDolar : acc.netEnding;

            return (
              <div 
                key={acc.code} 
                className="enterprise-card rounded-2xl overflow-hidden border border-slate-300 shadow-sm"
              >
                {/* Account Header Banner (Interactive on screen, formal accounting strip in print) */}
                <div 
                  onClick={() => toggleCollapse(acc.code)}
                  className="bg-slate-100 hover:bg-slate-200/80 p-3.5 print:p-2 print:bg-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-3 cursor-pointer select-none transition border-b-2 border-slate-300 print:border-slate-400"
                >
                  <div className="flex items-center gap-2.5">
                    <span className="p-1 rounded-md bg-white border border-slate-300 text-slate-600 shadow-2xs no-print">
                      {isCollapsed ? <ChevronLeft className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                    </span>
                    <span className="font-mono bg-blue-700 print:bg-slate-900 text-white px-2.5 py-0.5 rounded text-xs font-bold shadow-2xs">
                      {acc.code}
                    </span>
                    <div>
                      <div className="flex items-center gap-2">
                        <span className="font-sans font-black text-sm text-slate-900">
                          {acc.name}
                        </span>
                        <span className="bg-slate-200 text-slate-800 text-[10px] px-2 py-0.5 rounded font-bold">
                          {acc.category}
                        </span>
                        {acc.parentName && (
                          <span className="text-[11px] text-slate-500 font-sans hidden sm:inline">
                            (الأب: {acc.parentName})
                          </span>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Summary Badges on Header */}
                  <div className="flex items-center gap-2.5 flex-wrap print:gap-3 text-xs">
                    <span className="font-mono text-slate-700 bg-slate-200/80 print:bg-transparent px-2 py-0.5 rounded text-[11px] font-bold">
                      {mode === 'detailed' 
                        ? `${acc.transactionsCount || 0} حركة قيد` 
                        : `${acc.childAccountsCount || 0} حساب فرعي`}
                    </span>

                    <div className="text-left bg-white print:bg-transparent px-2.5 py-1 rounded-lg border border-slate-200 print:border-none flex items-center gap-1.5 shadow-2xs print:shadow-none">
                      <span className="text-[10px] text-slate-500 font-sans">رصيد سابق:</span>
                      {renderNetBadge(netOpenVal, false, currency === 'USD' ? '$' : 'ج.س')}
                    </div>

                    <div className="text-left bg-white print:bg-transparent px-2.5 py-1 rounded-lg border border-slate-200 print:border-none flex items-center gap-1.5 shadow-2xs print:shadow-none">
                      <span className="text-[10px] text-slate-500 font-sans">رصيد ختامي:</span>
                      {renderNetBadge(netEndingVal, true, currency === 'USD' ? '$' : 'ج.س')}
                    </div>
                  </div>
                </div>

                {/* Account Body Content (Collapsed by default on screen, printed fully on print) */}
                <div className={`${isCollapsed ? 'hidden print:block' : 'block'} overflow-x-auto`}>
                    {mode === 'detailed' ? (
                      /* DETAILED GROUP BY SUB-ACCOUNT (الحساب التحليلي) */
                      <div className="divide-y divide-slate-200">
                        {/* Sub-account actions toolbar */}
                        <div className="flex flex-wrap items-center justify-between gap-2 px-4 py-2.5 bg-slate-100/80 border-b border-slate-200 text-xs no-print">
                          <div className="flex items-center gap-2 text-slate-700 font-medium">
                            <FolderTree className="w-4 h-4 text-blue-600" />
                            <span>الحسابات التحليلية المجمعة:</span>
                            <span className="font-bold font-mono bg-blue-100 text-blue-800 px-2.5 py-0.5 rounded-full text-xs">
                              {acc.subAccounts?.length || 0} حساب تحليلي
                            </span>
                          </div>
                          <div className="flex items-center gap-2">
                            <button
                              type="button"
                              onClick={() => expandAllSubAccounts(acc)}
                              className="text-blue-700 hover:text-blue-900 font-semibold text-[11px] px-2.5 py-1 bg-white hover:bg-blue-50 border border-slate-200 rounded-lg shadow-2xs transition"
                            >
                              توسيع كافة الحسابات
                            </button>
                            <button
                              type="button"
                              onClick={() => collapseAllSubAccounts(acc)}
                              className="text-slate-600 hover:text-slate-800 font-semibold text-[11px] px-2.5 py-1 bg-white hover:bg-slate-50 border border-slate-200 rounded-lg shadow-2xs transition"
                            >
                              طي كافة الحسابات
                            </button>
                          </div>
                        </div>

                        {/* List of Sub-Accounts */}
                        {!acc.subAccounts || acc.subAccounts.length === 0 ? (
                          <div className="py-8 text-center text-slate-400 font-sans text-xs">
                            لا توجد قيود أو حسابات فرعية مسجلة لهذا الحساب خلال الفترة المحددة
                          </div>
                        ) : (
                          acc.subAccounts.map((sub) => {
                            const isSubCollapsed = !expandedSubAccounts[sub.accountNumber];
                            return (
                              <div key={sub.accountNumber} className="bg-white border-b last:border-b-0 border-slate-200">
                                {/* Sub-account Header Banner */}
                                <div
                                  onClick={() => toggleSubAccount(sub.accountNumber)}
                                  className="flex items-center justify-between gap-2 px-3 py-2 print:px-1.5 print:py-1 bg-gradient-to-r from-slate-50 to-blue-50/40 hover:bg-blue-50/70 cursor-pointer border-b border-slate-300 transition text-xs print:text-[10px] print:bg-slate-100"
                                >
                                  <div className="flex items-center gap-1.5 shrink-0 overflow-hidden">
                                    <button
                                      type="button"
                                      className="p-0.5 rounded text-slate-500 hover:text-slate-800 hover:bg-slate-200/60 no-print"
                                      aria-label="تبديل العرض"
                                    >
                                      {isSubCollapsed ? <ChevronLeft className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                                    </button>
                                    <span className="font-mono text-xs print:text-[9.5px] font-bold bg-white text-blue-800 px-1.5 py-0.5 rounded border border-blue-200">
                                      {sub.accountNumber}
                                    </span>
                                    <span className="font-bold text-slate-900 text-xs print:text-[10.5px] truncate">
                                      {sub.accountName}
                                    </span>
                                    <span className="text-[10px] font-sans font-medium bg-slate-200/80 text-slate-700 px-1.5 py-0.2 rounded-full">
                                      {sub.transactions.length} قيد
                                    </span>
                                  </div>

                                  {/* Balances badges on the sub-account header */}
                                  <div className="flex items-center gap-2 print:gap-1.5 text-xs print:text-[9.5px] font-mono shrink-0">
                                    <div className="flex items-center gap-1">
                                      <span className="text-slate-500 font-sans text-[10px]">رصيد سابق:</span>
                                      {(currency === 'EGP' || currency === 'BOTH') && renderNetBadge(sub.netOpen, false, 'ج.س')}
                                      {(currency === 'USD' || currency === 'BOTH') && renderNetBadge(sub.netOpenDolar, false, '$')}
                                    </div>
                                    <div className="flex items-center gap-1">
                                      <span className="text-slate-500 font-sans text-[10px]">صافي ختامي:</span>
                                      {(currency === 'EGP' || currency === 'BOTH') && renderNetBadge(sub.netEnding, true, 'ج.س')}
                                      {(currency === 'USD' || currency === 'BOTH') && renderNetBadge(sub.netEndingDolar, true, '$')}
                                    </div>
                                  </div>
                                </div>

                                {/* Sub-account Transactions Table (Shown when not collapsed, always shown on print) */}
                                <div className={`${isSubCollapsed ? 'hidden print:block' : 'block'} overflow-x-auto print:overflow-visible w-full`}>
                                  <table className="w-full text-right text-xs print:text-[9px] border-collapse table-fixed">
                                    <thead className="bg-slate-100 text-slate-800 font-bold border-b border-slate-300">
                                      <tr>
                                        <th className="py-1.5 px-2 print:px-1 border-l border-slate-300 w-[11%] text-center">التاريخ</th>
                                        <th className="py-1.5 px-2 print:px-1 border-l border-slate-300 w-[13%] text-center">رقم القيد / السند</th>
                                        <th className="py-1.5 px-2 print:px-1 border-l border-slate-300 w-[9%] text-center">نوع السند</th>
                                        <th className="py-1.5 px-2 print:px-1 border-l border-slate-300 w-[31%] text-right">البيان والشرح</th>
                                        
                                        {(currency === 'EGP' || currency === 'BOTH') && (
                                          <>
                                            <th className="py-1.5 px-2 print:px-1 text-left font-mono text-rose-800 bg-rose-50/50 w-[12%]">مدين (ج.س)</th>
                                            <th className="py-1.5 px-2 print:px-1 text-left font-mono text-emerald-800 bg-emerald-50/50 border-l border-slate-300 w-[12%]">دائن (ج.س)</th>
                                          </>
                                        )}

                                        {(currency === 'USD' || currency === 'BOTH') && (
                                          <>
                                            <th className="py-1.5 px-2 print:px-1 text-left font-mono text-rose-800 bg-rose-50/50 w-[12%]">مدين ($)</th>
                                            <th className="py-1.5 px-2 print:px-1 text-left font-mono text-emerald-800 bg-emerald-50/50 border-l border-slate-300 w-[12%]">دائن ($)</th>
                                          </>
                                        )}

                                        {(currency === 'EGP' || currency === 'BOTH') && (
                                          <th className="py-1.5 px-2 print:px-1 text-left font-mono bg-slate-100 border-l border-slate-300 w-[12%]">
                                            الرصيد التراكمي
                                          </th>
                                        )}

                                        {(currency === 'USD' || currency === 'BOTH') && (
                                          <th className="py-1.5 px-2 print:px-1 text-left font-mono bg-blue-50/60 w-[12%]">
                                            الرصيد التراكمي ($)
                                          </th>
                                        )}
                                      </tr>
                                    </thead>

                                    <tbody className="divide-y divide-slate-100 font-mono">
                                      {/* Row 0: Opening Balance Row for Sub-Account */}
                                      <tr className="bg-slate-50 font-bold border-b border-slate-200">
                                        <td className="py-1.5 px-2 print:px-1 text-slate-500 text-center w-[11%]">{startDate}</td>
                                        <td className="py-1.5 px-2 print:px-1 text-slate-400 font-sans text-center w-[13%]">-</td>
                                        <td className="py-1.5 px-2 print:px-1 text-blue-700 font-sans text-center w-[9%]">
                                          <span className="bg-blue-100 text-blue-800 px-1.5 py-0.2 rounded text-[10px]">رصيد سابق</span>
                                        </td>
                                        <td className="py-1.5 px-2 print:px-1 font-sans text-slate-600 font-medium w-[31%] truncate">
                                          الرصيد الافتتاحي السابق المدور لحساب [{sub.accountName}] عن فترة ما قبل {startDate}
                                        </td>

                                        {(currency === 'EGP' || currency === 'BOTH') && (
                                          <>
                                            <td className="py-1.5 px-2 print:px-1 text-left text-slate-300 w-[12%]">-</td>
                                            <td className="py-1.5 px-2 print:px-1 text-left text-slate-300 border-l border-slate-200 w-[12%]">-</td>
                                          </>
                                        )}
                                        {(currency === 'USD' || currency === 'BOTH') && (
                                          <>
                                            <td className="py-1.5 px-2 print:px-1 text-left text-slate-300 w-[12%]">-</td>
                                            <td className="py-1.5 px-2 print:px-1 text-left text-slate-300 border-l border-slate-200 w-[12%]">-</td>
                                          </>
                                        )}

                                        {(currency === 'EGP' || currency === 'BOTH') && (
                                          <td className="py-1.5 px-2 print:px-1 text-left bg-slate-100/70 border-l border-slate-200 w-[12%]">
                                            <div className="flex items-center justify-between gap-1 w-full font-mono text-xs print:text-[9px]">
                                              <span className="font-bold text-slate-900 tabular-nums">
                                                {sub.netOpen !== 0 ? formatCurrency(Math.abs(sub.netOpen)) : '-'}
                                              </span>
                                              {sub.netOpen !== 0 && (
                                                <span className={`text-[8.5px] font-sans font-bold px-1 py-0.2 rounded shrink-0 ${
                                                  sub.netOpen > 0 ? 'bg-rose-100 text-rose-800' : 'bg-emerald-100 text-emerald-800'
                                                }`}>
                                                  {sub.netOpen > 0 ? 'مدين' : 'دائن'}
                                                </span>
                                              )}
                                            </div>
                                          </td>
                                        )}
                                        {(currency === 'USD' || currency === 'BOTH') && (
                                          <td className="py-1.5 px-2 print:px-1 text-left bg-blue-100/40 w-[12%]">
                                            <div className="flex items-center justify-between gap-1 w-full font-mono text-xs print:text-[9px]">
                                              <span className="font-bold text-slate-900 tabular-nums">
                                                {sub.netOpenDolar !== 0 ? `$${formatCurrency(Math.abs(sub.netOpenDolar))}` : '-'}
                                              </span>
                                              {sub.netOpenDolar !== 0 && (
                                                <span className={`text-[8.5px] font-sans font-bold px-1 py-0.2 rounded shrink-0 ${
                                                  sub.netOpenDolar > 0 ? 'bg-rose-100 text-rose-800' : 'bg-emerald-100 text-emerald-800'
                                                }`}>
                                                  {sub.netOpenDolar > 0 ? 'مدين' : 'دائن'}
                                                </span>
                                              )}
                                            </div>
                                          </td>
                                        )}
                                      </tr>

                                      {/* Transaction Rows for Sub-Account */}
                                      {sub.transactions.length === 0 ? (
                                        <tr>
                                          <td colSpan={10} className="py-4 text-center text-slate-400 font-sans">
                                            لا توجد حركات خلال الفترة لهذا الحساب (الرصيد منقول فقط)
                                          </td>
                                        </tr>
                                      ) : (
                                        sub.transactions.map((t, tIdx) => (
                                          <tr key={`subtx-${t.Id}-${tIdx}`} className="hover:bg-blue-50/40 transition">
                                            <td className="py-1.5 px-2 print:px-1 text-slate-600 text-center w-[11%]">
                                              {formatDate(t.TheDate)}
                                            </td>
                                            <td className="py-1.5 px-2 print:px-1 text-center w-[13%]">
                                              <button
                                                type="button"
                                                onClick={() => openVoucherModal(t.ReqNum)}
                                                className="font-bold text-blue-600 hover:text-blue-800 flex items-center justify-center gap-1 cursor-pointer w-full"
                                                title="انقر لاستعراض سند القيد المحاسبي كاملاً"
                                              >
                                                <span>{t.ReqNum}</span>
                                                <ExternalLink className="w-3 h-3 opacity-60 no-print" />
                                              </button>
                                            </td>
                                            <td className="py-1.5 px-2 print:px-1 font-sans text-center w-[9%]">
                                              <span className="bg-slate-100 text-slate-700 px-1.5 py-0.2 rounded text-[10px] font-semibold border border-slate-200">
                                                {t.DocType || 'قيد'}
                                              </span>
                                            </td>
                                            <td className="py-1.5 px-2 print:px-1 font-sans text-slate-700 text-right w-[31%] break-words leading-tight" title={t.Detail}>
                                              {t.Detail || '-'}
                                            </td>

                                            {/* Debtor & Creditor EGP */}
                                            {(currency === 'EGP' || currency === 'BOTH') && (
                                              <>
                                                <td className="py-1.5 px-2 print:px-1 text-left font-bold text-rose-700 bg-rose-50/20 w-[12%] tabular-nums">
                                                  {t.Debtor > 0 ? formatCurrency(t.Debtor) : <span className="text-slate-300">-</span>}
                                                </td>
                                                <td className="py-1.5 px-2 print:px-1 text-left font-bold text-emerald-700 bg-emerald-50/20 border-l border-slate-200 w-[12%] tabular-nums">
                                                  {t.Creditor > 0 ? formatCurrency(t.Creditor) : <span className="text-slate-300">-</span>}
                                                </td>
                                              </>
                                            )}

                                            {/* Debtor & Creditor USD */}
                                            {(currency === 'USD' || currency === 'BOTH') && (
                                              <>
                                                <td className="py-1.5 px-2 print:px-1 text-left font-bold text-rose-700 bg-rose-50/20 w-[12%] tabular-nums">
                                                  {t.DebtorDolar > 0 ? `$${formatCurrency(t.DebtorDolar)}` : <span className="text-slate-300">-</span>}
                                                </td>
                                                <td className="py-1.5 px-2 print:px-1 text-left font-bold text-emerald-700 bg-emerald-50/20 border-l border-slate-200 w-[12%] tabular-nums">
                                                  {t.CreditorDolar > 0 ? `$${formatCurrency(t.CreditorDolar)}` : <span className="text-slate-300">-</span>}
                                                </td>
                                              </>
                                            )}

                                            {/* Running Balances for this Sub-Account */}
                                            {(currency === 'EGP' || currency === 'BOTH') && (
                                              <td className="py-1.5 px-2 print:px-1 text-left bg-slate-100/40 border-l border-slate-200 w-[12%]">
                                                <div className="flex items-center justify-between gap-1 w-full font-mono text-xs print:text-[9px]">
                                                  <span className="font-bold text-slate-900 tabular-nums">
                                                    {t.runningNet !== 0 ? formatCurrency(Math.abs(t.runningNet)) : '-'}
                                                  </span>
                                                  {t.runningNet !== 0 && (
                                                    <span className={`text-[8.5px] font-sans font-bold px-1 py-0.2 rounded shrink-0 ${
                                                      t.runningNet > 0 ? 'bg-rose-100 text-rose-800' : 'bg-emerald-100 text-emerald-800'
                                                    }`}>
                                                      {t.runningNet > 0 ? 'مدين' : 'دائن'}
                                                    </span>
                                                  )}
                                                </div>
                                              </td>
                                            )}
                                            {(currency === 'USD' || currency === 'BOTH') && (
                                              <td className="py-1.5 px-2 print:px-1 text-left bg-blue-50/30 w-[12%]">
                                                <div className="flex items-center justify-between gap-1 w-full font-mono text-xs print:text-[9px]">
                                                  <span className="font-bold text-slate-900 tabular-nums">
                                                    {t.runningNetDolar !== 0 ? `$${formatCurrency(Math.abs(t.runningNetDolar))}` : '-'}
                                                  </span>
                                                  {t.runningNetDolar !== 0 && (
                                                    <span className={`text-[8.5px] font-sans font-bold px-1 py-0.2 rounded shrink-0 ${
                                                      t.runningNetDolar > 0 ? 'bg-rose-100 text-rose-800' : 'bg-emerald-100 text-emerald-800'
                                                    }`}>
                                                      {t.runningNetDolar > 0 ? 'مدين' : 'دائن'}
                                                    </span>
                                                  )}
                                                </div>
                                              </td>
                                            )}
                                          </tr>
                                        ))
                                      )}
                                    </tbody>

                                    {/* Sub-account Subtotal Row */}
                                    <tfoot className="bg-slate-100 font-bold border-t border-slate-300 text-xs print:text-[9px]">
                                      <tr>
                                        <td colSpan={4} className="py-1.5 px-2 print:px-1 font-sans text-slate-800 w-[64%]">
                                          مجموع حركات وصافي رصيد [{sub.accountName}]:
                                        </td>

                                        {(currency === 'EGP' || currency === 'BOTH') && (
                                          <>
                                            <td className="py-1.5 px-2 print:px-1 text-left text-rose-800 bg-rose-50/70 w-[12%] tabular-nums">
                                              {sub.periodDebtor > 0 ? formatCurrency(sub.periodDebtor) : '-'}
                                            </td>
                                            <td className="py-1.5 px-2 print:px-1 text-left text-emerald-800 bg-emerald-50/70 border-l border-slate-200 w-[12%] tabular-nums">
                                              {sub.periodCreditor > 0 ? formatCurrency(sub.periodCreditor) : '-'}
                                            </td>
                                          </>
                                        )}

                                        {(currency === 'USD' || currency === 'BOTH') && (
                                          <>
                                            <td className="py-1.5 px-2 print:px-1 text-left text-rose-800 bg-rose-50/70 w-[12%] tabular-nums">
                                              {sub.periodDebtorDolar > 0 ? `$${formatCurrency(sub.periodDebtorDolar)}` : '-'}
                                            </td>
                                            <td className="py-1.5 px-2 print:px-1 text-left text-emerald-800 bg-emerald-50/70 border-l border-slate-200 w-[12%] tabular-nums">
                                              {sub.periodCreditorDolar > 0 ? `$${formatCurrency(sub.periodCreditorDolar)}` : '-'}
                                            </td>
                                          </>
                                        )}

                                        {(currency === 'EGP' || currency === 'BOTH') && (
                                          <td className="py-1.5 px-2 print:px-1 text-left bg-slate-200/60 border-l border-slate-200 w-[12%]">
                                            <div className="flex items-center justify-between gap-1 w-full font-mono text-xs print:text-[9px]">
                                              <span className="font-bold text-slate-950 tabular-nums">
                                                {sub.netEnding !== 0 ? formatCurrency(Math.abs(sub.netEnding)) : '-'}
                                              </span>
                                              {sub.netEnding !== 0 && (
                                                <span className={`text-[8.5px] font-sans font-bold px-1 py-0.2 rounded shrink-0 ${
                                                  sub.netEnding > 0 ? 'bg-rose-100 text-rose-800' : 'bg-emerald-100 text-emerald-800'
                                                }`}>
                                                  {sub.netEnding > 0 ? 'مدين' : 'دائن'}
                                                </span>
                                              )}
                                            </div>
                                          </td>
                                        )}

                                        {(currency === 'USD' || currency === 'BOTH') && (
                                          <td className="py-1.5 px-2 print:px-1 text-left bg-blue-100/50 w-[12%]">
                                            <div className="flex items-center justify-between gap-1 w-full font-mono text-xs print:text-[9px]">
                                              <span className="font-bold text-slate-950 tabular-nums">
                                                {sub.netEndingDolar !== 0 ? `$${formatCurrency(Math.abs(sub.netEndingDolar))}` : '-'}
                                              </span>
                                              {sub.netEndingDolar !== 0 && (
                                                <span className={`text-[8.5px] font-sans font-bold px-1 py-0.2 rounded shrink-0 ${
                                                  sub.netEndingDolar > 0 ? 'bg-rose-100 text-rose-800' : 'bg-emerald-100 text-emerald-800'
                                                }`}>
                                                  {sub.netEndingDolar > 0 ? 'مدين' : 'دائن'}
                                                </span>
                                              )}
                                            </div>
                                          </td>
                                        )}
                                      </tr>
                                    </tfoot>
                                  </table>
                                </div>
                              </div>
                            );
                          })
                        )}

                        {/* Grand Total Row for the entire General Account */}
                        <div className="bg-slate-200/90 p-2.5 print:p-1.5 border-t-2 border-slate-400 w-full">
                          <div className="flex items-center justify-between gap-2 text-xs print:text-[9.5px] font-bold w-full">
                            <div className="font-sans text-slate-900 text-xs print:text-[10px] truncate">
                              إجمالي الحساب العام بالكامل [{acc.code} - {acc.name}] ({acc.subAccounts?.length || 0} حساب تحليلي):
                            </div>

                            <div className="flex items-center gap-2 print:gap-1.5 font-mono shrink-0 text-xs print:text-[9.5px]">
                              {(currency === 'EGP' || currency === 'BOTH') && (
                                <div className="flex items-center gap-2">
                                  <span className="text-slate-600 font-sans text-xs">حركات الفترة (ج.س):</span>
                                  <span className="text-rose-800">مدين: {formatCurrency(acc.periodDebtor)}</span>
                                  <span>/</span>
                                  <span className="text-emerald-800">دائن: {formatCurrency(acc.periodCreditor)}</span>
                                  <span className="mr-2">الصافي الختامي:</span>
                                  {renderNetBadge(acc.netEnding, true, 'ج.س')}
                                </div>
                              )}

                              {(currency === 'USD' || currency === 'BOTH') && (
                                <div className="flex items-center gap-2">
                                  <span className="text-slate-600 font-sans text-xs">حركات ($):</span>
                                  <span className="text-blue-800">مدين: ${formatCurrency(acc.periodDebtorDolar)}</span>
                                  <span>/</span>
                                  <span className="text-blue-800">دائن: ${formatCurrency(acc.periodCreditorDolar)}</span>
                                  <span className="mr-2">الصافي الختامي:</span>
                                  {renderNetBadge(acc.netEndingDolar, true, '$')}
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      </div>
                    ) : (
                      /* SUMMARY CHILD ACCOUNTS TABLE */
                      <table className="w-full text-right text-xs border-collapse">
                        <thead className="bg-slate-100 text-slate-800 font-bold border-b-2 border-slate-300">
                          <tr>
                            <th className="py-2 px-2.5 border-l border-slate-300 font-mono w-[14%] whitespace-nowrap">كود الحساب الفرعي</th>
                            <th className="py-2 px-2.5 border-l border-slate-300 w-[36%]">اسم الحساب التحليلي</th>
                            
                            {(currency === 'EGP' || currency === 'BOTH') && (
                              <th className="py-2 px-2.5 text-left font-mono bg-slate-100 border-l border-slate-300 w-[14%] whitespace-nowrap">
                                الرصيد الافتتاحي (ج.س)
                              </th>
                            )}
                            {(currency === 'USD' || currency === 'BOTH') && (
                              <th className="py-2 px-2.5 text-left font-mono bg-blue-50/70 border-l border-slate-300 w-[14%] whitespace-nowrap">
                                الرصيد الافتتاحي ($)
                              </th>
                            )}

                            {(currency === 'EGP' || currency === 'BOTH') && (
                              <>
                                <th className="py-2 px-2.5 text-left font-mono text-rose-800 bg-rose-50/60 w-[11%] whitespace-nowrap">حركة مدين (ج.س)</th>
                                <th className="py-2 px-2.5 text-left font-mono text-emerald-800 bg-emerald-50/60 border-l border-slate-300 w-[11%] whitespace-nowrap">حركة دائن (ج.س)</th>
                              </>
                            )}
                            {(currency === 'USD' || currency === 'BOTH') && (
                              <>
                                <th className="py-2 px-2.5 text-left font-mono text-rose-800 bg-rose-50/60 w-[11%] whitespace-nowrap">حركة مدين ($)</th>
                                <th className="py-2 px-2.5 text-left font-mono text-emerald-800 bg-emerald-50/60 border-l border-slate-300 w-[11%] whitespace-nowrap">حركة دائن ($)</th>
                              </>
                            )}

                            {(currency === 'EGP' || currency === 'BOTH') && (
                              <th className="py-2 px-2.5 text-left font-mono bg-slate-100 border-l border-slate-300 w-[14%] whitespace-nowrap">
                                الرصيد الختامي (ج.س)
                              </th>
                            )}
                            {(currency === 'USD' || currency === 'BOTH') && (
                              <th className="py-2 px-2.5 text-left font-mono bg-blue-50/70 w-[14%] whitespace-nowrap">
                                الرصيد الختامي ($)
                              </th>
                            )}
                          </tr>
                        </thead>

                        <tbody className="divide-y divide-slate-200 font-mono">
                          {!acc.childAccounts || acc.childAccounts.length === 0 ? (
                            <tr>
                              <td colSpan={10} className="py-8 text-center text-slate-400 font-sans">
                                لا توجد حسابات فرعية تحليلية مسجلة تحت هذا الحساب العام
                              </td>
                            </tr>
                          ) : (
                            acc.childAccounts.map((c, cIdx) => (
                              <tr key={`child-${c.AccountNumber}-${cIdx}`} className="hover:bg-blue-50/40 transition border-b border-slate-200">
                                <td className="py-1.5 px-2.5 font-bold text-blue-700 border-l border-slate-200 whitespace-nowrap">
                                  {c.AccountNumber}
                                </td>
                                <td className="py-1.5 px-2.5 font-sans font-bold text-slate-800 border-l border-slate-200">
                                  {c.AccountName}
                                </td>

                                {/* Openings */}
                                {(currency === 'EGP' || currency === 'BOTH') && (
                                  <td className="py-1.5 px-2.5 text-left bg-slate-50/30 border-l border-slate-200 whitespace-nowrap">
                                    <div className="flex items-center justify-between gap-1 w-full font-mono text-xs">
                                      <span className="font-bold text-slate-900 tabular-nums">
                                        {c.netOpen !== 0 ? formatCurrency(Math.abs(c.netOpen)) : '-'}
                                      </span>
                                      {c.netOpen !== 0 && (
                                        <span className={`text-[10px] font-sans font-bold px-1.5 py-0.2 rounded shrink-0 ${
                                          c.netOpen > 0 
                                            ? 'bg-rose-50 text-rose-800 border border-rose-200' 
                                            : 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                        }`}>
                                          {c.netOpen > 0 ? 'مدين' : 'دائن'}
                                        </span>
                                      )}
                                    </div>
                                  </td>
                                )}
                                {(currency === 'USD' || currency === 'BOTH') && (
                                  <td className="py-1.5 px-2.5 text-left bg-blue-50/20 border-l border-slate-200 whitespace-nowrap">
                                    <div className="flex items-center justify-between gap-1 w-full font-mono text-xs">
                                      <span className="font-bold text-slate-900 tabular-nums">
                                        {c.netOpenDolar !== 0 ? `$${formatCurrency(Math.abs(c.netOpenDolar))}` : '-'}
                                      </span>
                                      {c.netOpenDolar !== 0 && (
                                        <span className={`text-[10px] font-sans font-bold px-1.5 py-0.2 rounded shrink-0 ${
                                          c.netOpenDolar > 0 
                                            ? 'bg-rose-50 text-rose-800 border border-rose-200' 
                                            : 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                        }`}>
                                          {c.netOpenDolar > 0 ? 'مدين' : 'دائن'}
                                        </span>
                                      )}
                                    </div>
                                  </td>
                                )}

                                {/* Period Movements */}
                                {(currency === 'EGP' || currency === 'BOTH') && (
                                  <>
                                    <td className="py-1.5 px-2.5 text-left font-bold text-rose-700 bg-rose-50/10 whitespace-nowrap tabular-nums">
                                      {c.periodDebtor > 0 ? formatCurrency(c.periodDebtor) : <span className="text-slate-300">-</span>}
                                    </td>
                                    <td className="py-1.5 px-2.5 text-left font-bold text-emerald-700 bg-emerald-50/10 border-l border-slate-200 whitespace-nowrap tabular-nums">
                                      {c.periodCreditor > 0 ? formatCurrency(c.periodCreditor) : <span className="text-slate-300">-</span>}
                                    </td>
                                  </>
                                )}
                                {(currency === 'USD' || currency === 'BOTH') && (
                                  <>
                                    <td className="py-1.5 px-2.5 text-left font-bold text-rose-700 bg-rose-50/10 whitespace-nowrap tabular-nums">
                                      {c.periodDebtorDolar > 0 ? `$${formatCurrency(c.periodDebtorDolar)}` : <span className="text-slate-300">-</span>}
                                    </td>
                                    <td className="py-1.5 px-2.5 text-left font-bold text-emerald-700 bg-emerald-50/10 border-l border-slate-200 whitespace-nowrap tabular-nums">
                                      {c.periodCreditorDolar > 0 ? `$${formatCurrency(c.periodCreditorDolar)}` : <span className="text-slate-300">-</span>}
                                    </td>
                                  </>
                                )}

                                {/* Endings */}
                                {(currency === 'EGP' || currency === 'BOTH') && (
                                  <td className="py-1.5 px-2.5 text-left bg-slate-100/40 border-l border-slate-200 whitespace-nowrap">
                                    <div className="flex items-center justify-between gap-1 w-full font-mono text-xs">
                                      <span className="font-extrabold text-slate-950 tabular-nums">
                                        {c.netEnding !== 0 ? formatCurrency(Math.abs(c.netEnding)) : '-'}
                                      </span>
                                      {c.netEnding !== 0 && (
                                        <span className={`text-[10px] font-sans font-bold px-1.5 py-0.2 rounded shrink-0 ${
                                          c.netEnding > 0 
                                            ? 'bg-rose-50 text-rose-800 border border-rose-200' 
                                            : 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                        }`}>
                                          {c.netEnding > 0 ? 'مدين' : 'دائن'}
                                        </span>
                                      )}
                                    </div>
                                  </td>
                                )}
                                {(currency === 'USD' || currency === 'BOTH') && (
                                  <td className="py-1.5 px-2.5 text-left bg-blue-50/30 whitespace-nowrap">
                                    <div className="flex items-center justify-between gap-1 w-full font-mono text-xs">
                                      <span className="font-extrabold text-slate-950 tabular-nums">
                                        {c.netEndingDolar !== 0 ? `$${formatCurrency(Math.abs(c.netEndingDolar))}` : '-'}
                                      </span>
                                      {c.netEndingDolar !== 0 && (
                                        <span className={`text-[10px] font-sans font-bold px-1.5 py-0.2 rounded shrink-0 ${
                                          c.netEndingDolar > 0 
                                            ? 'bg-rose-50 text-rose-800 border border-rose-200' 
                                            : 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                        }`}>
                                          {c.netEndingDolar > 0 ? 'مدين' : 'دائن'}
                                        </span>
                                      )}
                                    </div>
                                  </td>
                                )}
                              </tr>
                            ))
                          )}
                        </tbody>

                        {/* Summary Subtotal Row */}
                        <tfoot className="bg-slate-100 font-bold border-t-2 border-b-2 border-slate-400 text-xs">
                          <tr>
                            <td colSpan={2} className="py-2 px-3 font-sans text-slate-900 border-l border-slate-300">
                              إجمالي صافي الحسابات الفرعية لحساب [{acc.code} - {acc.name}]:
                            </td>

                            {(currency === 'EGP' || currency === 'BOTH') && (
                              <td className="py-2 px-3 text-left bg-slate-200/50 border-l border-slate-300">
                                <div className="flex items-center justify-between gap-1 w-full font-mono text-xs">
                                  <span className="font-bold text-slate-950 tabular-nums">
                                    {acc.netOpen !== 0 ? formatCurrency(Math.abs(acc.netOpen)) : '-'}
                                  </span>
                                  {acc.netOpen !== 0 && (
                                    <span className={`text-[10px] font-sans font-bold px-1.5 py-0.2 rounded shrink-0 ${
                                      acc.netOpen > 0 ? 'bg-rose-100 text-rose-800 border border-rose-300' : 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                    }`}>
                                      {acc.netOpen > 0 ? 'مدين' : 'دائن'}
                                    </span>
                                  )}
                                </div>
                              </td>
                            )}
                            {(currency === 'USD' || currency === 'BOTH') && (
                              <td className="py-2 px-3 text-left bg-blue-100/40 border-l border-slate-300">
                                <div className="flex items-center justify-between gap-1 w-full font-mono text-xs">
                                  <span className="font-bold text-slate-950 tabular-nums">
                                    {acc.netOpenDolar !== 0 ? `$${formatCurrency(Math.abs(acc.netOpenDolar))}` : '-'}
                                  </span>
                                  {acc.netOpenDolar !== 0 && (
                                    <span className={`text-[10px] font-sans font-bold px-1.5 py-0.2 rounded shrink-0 ${
                                      acc.netOpenDolar > 0 ? 'bg-rose-100 text-rose-800 border border-rose-300' : 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                    }`}>
                                      {acc.netOpenDolar > 0 ? 'مدين' : 'دائن'}
                                    </span>
                                  )}
                                </div>
                              </td>
                            )}

                            {(currency === 'EGP' || currency === 'BOTH') && (
                              <>
                                <td className="py-2 px-3 text-left text-rose-800 bg-rose-100/40 tabular-nums">
                                  {acc.periodDebtor > 0 ? formatCurrency(acc.periodDebtor) : '-'}
                                </td>
                                <td className="py-2 px-3 text-left text-emerald-800 bg-emerald-100/40 border-l border-slate-300 tabular-nums">
                                  {acc.periodCreditor > 0 ? formatCurrency(acc.periodCreditor) : '-'}
                                </td>
                              </>
                            )}
                            {(currency === 'USD' || currency === 'BOTH') && (
                              <>
                                <td className="py-2 px-3 text-left text-rose-800 bg-rose-100/40 tabular-nums">
                                  {acc.periodDebtorDolar > 0 ? `$${formatCurrency(acc.periodDebtorDolar)}` : '-'}
                                </td>
                                <td className="py-2 px-3 text-left text-emerald-800 bg-emerald-100/40 border-l border-slate-300 tabular-nums">
                                  {acc.periodCreditorDolar > 0 ? `$${formatCurrency(acc.periodCreditorDolar)}` : '-'}
                                </td>
                              </>
                            )}

                            {(currency === 'EGP' || currency === 'BOTH') && (
                              <td className="py-2 px-3 text-left bg-slate-300/50 border-l border-slate-300">
                                <div className="flex items-center justify-between gap-1 w-full font-mono text-xs">
                                  <span className="font-extrabold text-slate-950 tabular-nums">
                                    {acc.netEnding !== 0 ? formatCurrency(Math.abs(acc.netEnding)) : '-'}
                                  </span>
                                  {acc.netEnding !== 0 && (
                                    <span className={`text-[10px] font-sans font-bold px-1.5 py-0.2 rounded shrink-0 ${
                                      acc.netEnding > 0 ? 'bg-rose-100 text-rose-800 border border-rose-300' : 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                    }`}>
                                      {acc.netEnding > 0 ? 'مدين' : 'دائن'}
                                    </span>
                                  )}
                                </div>
                              </td>
                            )}
                            {(currency === 'USD' || currency === 'BOTH') && (
                              <td className="py-2 px-3 text-left bg-blue-100/60">
                                <div className="flex items-center justify-between gap-1 w-full font-mono text-xs">
                                  <span className="font-extrabold text-slate-950 tabular-nums">
                                    {acc.netEndingDolar !== 0 ? `$${formatCurrency(Math.abs(acc.netEndingDolar))}` : '-'}
                                  </span>
                                  {acc.netEndingDolar !== 0 && (
                                    <span className={`text-[10px] font-sans font-bold px-1.5 py-0.2 rounded shrink-0 ${
                                      acc.netEndingDolar > 0 ? 'bg-rose-100 text-rose-800 border border-rose-300' : 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                    }`}>
                                      {acc.netEndingDolar > 0 ? 'مدين' : 'دائن'}
                                    </span>
                                  )}
                                </div>
                              </td>
                            )}
                          </tr>
                        </tfoot>
                      </table>
                    )}
                  </div>
              </div>
            );
          })}
        </div>
      )}

      {/* 6. OVERALL GRAND TOTAL CARD */}
      {data && gt && (() => {
        const netOpenVal = (gt.openDebtor || 0) - (gt.openCreditor || 0);
        const netOpenValDolar = (gt.openDebtorDolar || 0) - (gt.openCreditorDolar || 0);

        const netEndingVal = (gt.totalDebtor || 0) - (gt.totalCreditor || 0);
        const netEndingValDolar = (gt.totalDebtorDolar || 0) - (gt.totalCreditorDolar || 0);

        return (
          <div className="rounded-2xl border border-slate-300 bg-white shadow-md overflow-hidden transition-all">
            {/* Header Strip with Royal Navy Gradient */}
            <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 px-6 py-4 flex flex-col md:flex-row md:items-center justify-between gap-3 text-white">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-blue-600/30 border border-blue-400/30 flex items-center justify-center text-blue-300 shadow-inner">
                  <Scale className="w-5 h-5 text-emerald-400" />
                </div>
                <div>
                  <h3 className="font-bold text-base text-white flex items-center gap-2">
                    <span>المجموع العام الإجمالي لكافة الحسابات العامة المختارة</span>
                    <span className="bg-emerald-500/20 text-emerald-300 border border-emerald-400/30 text-[11px] px-2.5 py-0.5 rounded-full font-sans font-semibold">
                      معتمد ومطابق
                    </span>
                  </h3>
                  <p className="text-xs text-slate-300 mt-0.5 font-sans">
                    تجميع شامل لكافة حركات وأرصدة <span className="font-bold text-white font-mono">{data.accounts.length}</span> حساب عام مفحوص خلال الفترة المحددة.
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2 self-start md:self-auto font-mono text-xs">
                <span className="bg-white/10 px-3 py-1.5 rounded-xl border border-white/10 text-slate-200">
                  العملة المعتمدة: <span className="font-bold text-amber-300 font-sans">{currency === 'EGP' ? 'جنيه سوداني (ج.س)' : currency === 'USD' ? 'دولار ($)' : 'العملتان معاً'}</span>
                </span>
              </div>
            </div>

            {/* 4 Executive Metric Cards Grid */}
            <div className="p-5 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 bg-slate-50/50">
              
              {/* Metric 1: صافي الرصيد الافتتاحي السابق */}
              <div className="bg-white p-4 rounded-xl border border-slate-200/90 shadow-2xs hover:shadow-xs transition">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-xs font-bold text-slate-700">صافي الرصيد الافتتاحي السابق</span>
                  <div className="w-7 h-7 rounded-lg bg-amber-50 text-amber-600 border border-amber-200/60 flex items-center justify-center">
                    <Coins className="w-3.5 h-3.5" />
                  </div>
                </div>
                <div className="text-[11px] text-slate-400 mb-2 font-sans">
                  الرصيد المدور قبل {startDate}
                </div>
                <div className="space-y-1.5 pt-1.5 border-t border-slate-100">
                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <div className="text-sm">
                      {renderNetBadge(netOpenVal, false, 'ج.س')}
                    </div>
                  )}
                  {(currency === 'USD' || currency === 'BOTH') && (
                    <div className="text-sm">
                      {renderNetBadge(netOpenValDolar, false, '$')}
                    </div>
                  )}
                </div>
              </div>

              {/* Metric 2: إجمالي حركة الفترة (مدين) */}
              <div className="bg-white p-4 rounded-xl border border-slate-200/90 shadow-2xs hover:shadow-xs transition">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-xs font-bold text-slate-700">إجمالي حركة الفترة (مدين)</span>
                  <div className="w-7 h-7 rounded-lg bg-rose-50 text-rose-600 border border-rose-200/60 flex items-center justify-center">
                    <ArrowDownLeft className="w-3.5 h-3.5" />
                  </div>
                </div>
                <div className="text-[11px] text-slate-400 mb-2 font-sans">
                  مجموع قيود وسندات القبض والصرف المدينة
                </div>
                <div className="space-y-1 pt-1.5 border-t border-slate-100">
                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <div className="font-mono text-lg font-black text-rose-700">
                      {formatCurrency(gt.periodDebtor)} <span className="text-xs font-sans font-bold text-slate-500">ج.س</span>
                    </div>
                  )}
                  {(currency === 'USD' || currency === 'BOTH') && (
                    <div className="font-mono text-base font-bold text-rose-800">
                      ${formatCurrency(gt.periodDebtorDolar)}
                    </div>
                  )}
                </div>
              </div>

              {/* Metric 3: إجمالي حركة الفترة (دائن) */}
              <div className="bg-white p-4 rounded-xl border border-slate-200/90 shadow-2xs hover:shadow-xs transition">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-xs font-bold text-slate-700">إجمالي حركة الفترة (دائن)</span>
                  <div className="w-7 h-7 rounded-lg bg-emerald-50 text-emerald-600 border border-emerald-200/60 flex items-center justify-center">
                    <TrendingUp className="w-3.5 h-3.5" />
                  </div>
                </div>
                <div className="text-[11px] text-slate-400 mb-2 font-sans">
                  مجموع قيود وسندات الصرف والتحويل الدائنة
                </div>
                <div className="space-y-1 pt-1.5 border-t border-slate-100">
                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <div className="font-mono text-lg font-black text-emerald-700">
                      {formatCurrency(gt.periodCreditor)} <span className="text-xs font-sans font-bold text-slate-500">ج.س</span>
                    </div>
                  )}
                  {(currency === 'USD' || currency === 'BOTH') && (
                    <div className="font-mono text-base font-bold text-emerald-800">
                      ${formatCurrency(gt.periodCreditorDolar)}
                    </div>
                  )}
                </div>
              </div>

              {/* Metric 4: صافي الرصيد الختامي التراكمي */}
              <div className="bg-gradient-to-br from-blue-50/70 via-white to-indigo-50/50 p-4 rounded-xl border-2 border-blue-200 shadow-2xs hover:shadow-xs transition">
                <div className="flex items-center justify-between mb-2">
                  <span className="text-xs font-extrabold text-blue-950">صافي الرصيد الختامي التراكمي</span>
                  <div className="w-7 h-7 rounded-lg bg-blue-600 text-white flex items-center justify-center shadow-2xs">
                    <CheckCircle2 className="w-4 h-4" />
                  </div>
                </div>
                <div className="text-[11px] text-slate-500 mb-2 font-sans">
                  الرصيد النهائي في ختام {endDate}
                </div>
                <div className="space-y-1.5 pt-1.5 border-t border-blue-100">
                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <div className="text-sm font-extrabold">
                      {renderNetBadge(netEndingVal, true, 'ج.س')}
                    </div>
                  )}
                  {(currency === 'USD' || currency === 'BOTH') && (
                    <div className="text-sm font-extrabold">
                      {renderNetBadge(netEndingValDolar, true, '$')}
                    </div>
                  )}
                </div>
              </div>

            </div>
          </div>
        );
      })()}

      {/* 7. PRINT SIGNATURES BLOCK */}
      <div className="hidden print:grid grid-cols-3 gap-8 mt-12 pt-6 border-t border-slate-400 text-center text-xs">
        <div>
          <span className="font-bold block text-slate-900">إعداد المحاسب المسؤول</span>
          <span className="mt-12 block border-b border-dashed border-slate-400 w-44 mx-auto"></span>
        </div>
        <div>
          <span className="font-bold block text-slate-900">مراجعة المراجع المالي</span>
          <span className="mt-12 block border-b border-dashed border-slate-400 w-44 mx-auto"></span>
        </div>
        <div>
          <span className="font-bold block text-slate-900">اعتماد المدير المالي (CFO)</span>
          <span className="mt-12 block border-b border-dashed border-slate-400 w-44 mx-auto"></span>
        </div>
      </div>
    </div>
  );
}

