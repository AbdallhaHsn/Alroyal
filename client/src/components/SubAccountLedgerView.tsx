import React, { useState, useMemo, useRef, useEffect } from 'react';
import {
  FileText,
  Calendar,
  FileSpreadsheet,
  Printer,
  ChevronDown,
  Search,
  Filter,
  X,
  Coins,
  TrendingUp,
  TrendingDown,
  ArrowDownLeft,
  RefreshCw,
  FolderTree,
  ExternalLink,
  CheckCircle2,
  Building2,
  Hash,
  Tag
} from 'lucide-react';

export interface SubAccountItem {
  code: string;
  name: string;
  parentCode: string;
  parentName: string;
  category: string;
}

export interface SubAccountLedgerTransaction {
  Id: number;
  TheDate: string;
  ReqNum: string;
  DocType: string;
  Detail: string;
  Debtor: number;
  Creditor: number;
  DebtorDolar: number;
  CreditorDolar: number;
  Curruncy_Name: string;
  Check_Number: string;
  runningNet: number;
  runningNetDolar: number;
}

export interface SubAccountLedgerData {
  startDate: string;
  endDate: string;
  account: {
    code: string;
    name: string;
    parentCode?: string;
    parentName?: string;
    category?: string;
  };
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
  transactionsCount: number;
  transactions: SubAccountLedgerTransaction[];
}

interface SubAccountLedgerViewProps {
  subAccounts: SubAccountItem[];
  selectedAccountCode: string;
  setSelectedAccountCode: (code: string) => void;
  startDate: string;
  setStartDate: (date: string) => void;
  endDate: string;
  setEndDate: (date: string) => void;
  currency: 'EGP' | 'USD' | 'BOTH';
  setCurrency: (currency: 'EGP' | 'USD' | 'BOTH') => void;
  loading: boolean;
  data: SubAccountLedgerData | null;
  onRefresh: () => void;
  onExportExcel: () => void;
  formatCurrency: (val: number) => string;
  renderNetBadge: (val: number, isEnding?: boolean, unit?: string) => React.ReactNode;
  openVoucherModal: (reqNum: string) => void;
}

export const SubAccountLedgerView: React.FC<SubAccountLedgerViewProps> = ({
  subAccounts,
  selectedAccountCode,
  setSelectedAccountCode,
  startDate,
  setStartDate,
  endDate,
  setEndDate,
  currency,
  setCurrency,
  loading,
  data,
  onRefresh,
  onExportExcel,
  formatCurrency,
  renderNetBadge,
  openVoucherModal
}) => {
  const [comboboxOpen, setComboboxOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('all');
  const dropdownRef = useRef<HTMLDivElement>(null);

  // Close dropdown on outside click
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setComboboxOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Currently selected account metadata
  const currentAccount = useMemo(() => {
    if (!selectedAccountCode) return null;
    return subAccounts.find(a => a.code === selectedAccountCode) || {
      code: selectedAccountCode,
      name: data?.account?.name || selectedAccountCode,
      parentCode: data?.account?.parentCode || '',
      parentName: data?.account?.parentName || '',
      category: data?.account?.category || 'الأصول'
    };
  }, [subAccounts, selectedAccountCode, data]);

  // Filtered sub-accounts for combobox
  const filteredSubAccounts = useMemo(() => {
    let list = subAccounts;
    if (categoryFilter !== 'all') {
      list = list.filter(a => a.category === categoryFilter || a.code.startsWith(categoryFilter));
    }
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      list = list.filter(a =>
        a.code.toLowerCase().includes(q) ||
        a.name.toLowerCase().includes(q) ||
        (a.parentName && a.parentName.toLowerCase().includes(q))
      );
    }
    return list;
  }, [subAccounts, categoryFilter, searchQuery]);

  const formatDate = (dateStr: string) => {
    if (!dateStr) return '';
    try {
      return new Date(dateStr).toISOString().split('T')[0];
    } catch {
      return dateStr;
    }
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="space-y-4">
      {/* 1. FILTER & CONTROL PANEL */}
      <div className="enterprise-card rounded-2xl p-5 no-print space-y-4">
        {/* Top Header of Filter Bar */}
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-3">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-blue-50 text-blue-700 rounded-xl">
              <FileText className="w-5 h-5" />
            </div>
            <div>
              <h2 className="font-bold text-slate-900 text-sm sm:text-base">
                كشف حساب الحسابات الفرعية والتحليلية (Subsidiary Ledger)
              </h2>
              <p className="text-xs text-slate-500">
                كشف حساب تفصيلي بالقيود والحركات والرصيد التراكمي لكل عميل أو مورد أو حساب فرعي.
              </p>
            </div>
          </div>

          {/* Action buttons */}
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onRefresh}
              disabled={loading || !selectedAccountCode}
              className="flex items-center gap-1.5 px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition disabled:opacity-50 cursor-pointer shadow-2xs"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin' : ''}`} />
              <span>تحديث</span>
            </button>

            <button
              type="button"
              onClick={onExportExcel}
              disabled={loading || !data}
              className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold transition disabled:opacity-50 shadow-xs shadow-emerald-600/20 cursor-pointer"
            >
              <FileSpreadsheet className="w-4 h-4" />
              <span>تصدير إكسيل منسق</span>
            </button>

            <button
              type="button"
              onClick={handlePrint}
              disabled={loading || !data}
              className="flex items-center gap-1.5 px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold transition disabled:opacity-50 shadow-xs shadow-blue-600/20 cursor-pointer"
            >
              <Printer className="w-4 h-4" />
              <span>طباعة كشف الحساب</span>
            </button>
          </div>
        </div>

        {/* Filter Controls Row */}
        <div className="grid grid-cols-1 md:grid-cols-12 gap-3.5 items-end">
          {/* Sub-Account Selector (Combobox / Direct Input) */}
          <div className="md:col-span-5 relative" ref={dropdownRef}>
            <label className="block text-xs font-bold text-slate-700 mb-1.5">
              تحديد الحساب الفرعي (رقم الحساب أو بالاسم):
            </label>

            <div className="relative">
              <div className="flex items-center rounded-xl border border-slate-300 bg-white shadow-2xs overflow-hidden focus-within:ring-2 focus-within:ring-blue-500 focus-within:border-blue-500">
                <div className="p-2.5 text-slate-400">
                  <Search className="w-4 h-4" />
                </div>
                <input
                  type="text"
                  value={searchQuery || (currentAccount ? `${currentAccount.code} - ${currentAccount.name}` : '')}
                  onChange={(e) => {
                    setSearchQuery(e.target.value);
                    setComboboxOpen(true);
                    // Check if exact code match was typed
                    const typed = e.target.value.trim();
                    const exact = subAccounts.find(a => a.code === typed);
                    if (exact) {
                      setSelectedAccountCode(exact.code);
                    }
                  }}
                  onFocus={() => {
                    setSearchQuery('');
                    setComboboxOpen(true);
                  }}
                  placeholder="اكتب رقم الحساب مباشرة أو ابحث بالاسم..."
                  className="w-full text-xs font-sans text-slate-800 placeholder-slate-400 py-2.5 pr-1 pl-2 focus:outline-none"
                />
                {selectedAccountCode && (
                  <button
                    type="button"
                    onClick={() => {
                      setSelectedAccountCode('');
                      setSearchQuery('');
                    }}
                    className="p-1.5 text-slate-400 hover:text-slate-600"
                    title="مسح الحساب المختار"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => setComboboxOpen(!comboboxOpen)}
                  className="p-2.5 text-slate-500 hover:bg-slate-100 border-r border-slate-200 transition"
                  title="استعراض شجرة الحسابات الفرعية"
                >
                  <ChevronDown className={`w-4 h-4 transition-transform ${comboboxOpen ? 'rotate-180' : ''}`} />
                </button>
              </div>

              {/* Combobox Dropdown Menu */}
              {comboboxOpen && (
                <div className="absolute z-50 mt-1.5 w-full bg-white rounded-2xl border border-slate-200 shadow-xl max-h-80 flex flex-col overflow-hidden animate-in fade-in slide-in-from-top-2 duration-150">
                  {/* Category Filter Tabs */}
                  <div className="flex items-center gap-1 p-2 bg-slate-50 border-b border-slate-200 overflow-x-auto text-[11px]">
                    <button
                      type="button"
                      onClick={() => setCategoryFilter('all')}
                      className={`px-2.5 py-1 rounded-lg font-bold transition whitespace-nowrap ${
                        categoryFilter === 'all' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      الكل ({subAccounts.length})
                    </button>
                    <button
                      type="button"
                      onClick={() => setCategoryFilter('01')}
                      className={`px-2.5 py-1 rounded-lg font-bold transition whitespace-nowrap ${
                        categoryFilter === '01' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      الأصول
                    </button>
                    <button
                      type="button"
                      onClick={() => setCategoryFilter('02')}
                      className={`px-2.5 py-1 rounded-lg font-bold transition whitespace-nowrap ${
                        categoryFilter === '02' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      الخصوم
                    </button>
                    <button
                      type="button"
                      onClick={() => setCategoryFilter('03')}
                      className={`px-2.5 py-1 rounded-lg font-bold transition whitespace-nowrap ${
                        categoryFilter === '03' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      المصروفات
                    </button>
                    <button
                      type="button"
                      onClick={() => setCategoryFilter('04')}
                      className={`px-2.5 py-1 rounded-lg font-bold transition whitespace-nowrap ${
                        categoryFilter === '04' ? 'bg-blue-600 text-white shadow-2xs' : 'text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      الإيرادات
                    </button>
                  </div>

                  {/* Accounts List */}
                  <div className="overflow-y-auto flex-1 divide-y divide-slate-100">
                    {filteredSubAccounts.length === 0 ? (
                      <div className="p-4 text-center text-slate-400 text-xs">
                        لا توجد نتائج مطابقة لبحثك
                      </div>
                    ) : (
                      filteredSubAccounts.map(acc => {
                        const isSelected = selectedAccountCode === acc.code;
                        return (
                          <div
                            key={acc.code}
                            onClick={() => {
                              setSelectedAccountCode(acc.code);
                              setSearchQuery('');
                              setComboboxOpen(false);
                            }}
                            className={`p-2.5 hover:bg-blue-50/70 cursor-pointer transition flex items-center justify-between text-xs ${
                              isSelected ? 'bg-blue-50 text-blue-900 font-bold' : 'text-slate-800'
                            }`}
                          >
                            <div className="flex items-center gap-2">
                              <span className="font-mono text-xs font-bold text-blue-700 bg-blue-100/70 px-2 py-0.5 rounded border border-blue-200/80">
                                {acc.code}
                              </span>
                              <span className="font-semibold">{acc.name}</span>
                            </div>
                            <div className="flex items-center gap-2 text-[11px] text-slate-500">
                              {acc.parentName && (
                                <span className="bg-slate-100 px-1.5 py-0.5 rounded text-slate-600">
                                  {acc.parentName}
                                </span>
                              )}
                              <span className="text-slate-400 font-mono">{acc.category}</span>
                            </div>
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Date Range Filters */}
          <div className="md:col-span-2">
            <label className="block text-xs font-bold text-slate-700 mb-1.5">
              من تاريخ:
            </label>
            <div className="relative">
              <input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="w-full text-xs font-mono rounded-xl border border-slate-300 py-2.5 px-3 bg-white text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>
          </div>

          <div className="md:col-span-2">
            <label className="block text-xs font-bold text-slate-700 mb-1.5">
              إلى تاريخ:
            </label>
            <div className="relative">
              <input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="w-full text-xs font-mono rounded-xl border border-slate-300 py-2.5 px-3 bg-white text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>
          </div>

          {/* Currency Selector */}
          <div className="md:col-span-3">
            <label className="block text-xs font-bold text-slate-700 mb-1.5">
              العملة المحاسبية:
            </label>
            <div className="grid grid-cols-3 gap-1 bg-slate-100 p-1 rounded-xl">
              <button
                type="button"
                onClick={() => setCurrency('EGP')}
                className={`py-1.5 text-xs font-bold rounded-lg transition ${
                  currency === 'EGP' ? 'bg-white text-blue-700 shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                جنيه سوداني (SDG)
              </button>
              <button
                type="button"
                onClick={() => setCurrency('USD')}
                className={`py-1.5 text-xs font-bold rounded-lg transition ${
                  currency === 'USD' ? 'bg-white text-blue-700 shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                دولار ($)
              </button>
              <button
                type="button"
                onClick={() => setCurrency('BOTH')}
                className={`py-1.5 text-xs font-bold rounded-lg transition ${
                  currency === 'BOTH' ? 'bg-white text-blue-700 shadow-2xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                كلاهما معاً
              </button>
            </div>
          </div>
        </div>

        {/* Selected Account Info Bar */}
        {currentAccount && (
          <div className="flex flex-wrap items-center justify-between gap-3 p-3 bg-blue-50/70 border border-blue-200/80 rounded-xl text-xs">
            <div className="flex items-center gap-2.5 flex-wrap">
              <span className="font-mono text-xs font-bold bg-blue-700 text-white px-2.5 py-1 rounded-lg">
                {currentAccount.code}
              </span>
              <span className="font-bold text-slate-900 text-sm">{currentAccount.name}</span>
              {currentAccount.parentName && (
                <span className="bg-white border border-blue-200 text-blue-800 px-2 py-0.5 rounded-full text-[11px] font-semibold">
                  الحساب العام: {currentAccount.parentName} ({currentAccount.parentCode})
                </span>
              )}
              <span className="bg-white border border-slate-200 text-slate-600 px-2 py-0.5 rounded-full text-[11px] font-medium">
                التصنيف: {currentAccount.category}
              </span>
            </div>

            <div className="text-[11px] text-slate-500 font-medium">
              الفترة: من <span className="font-mono text-slate-800 font-bold">{startDate}</span> إلى <span className="font-mono text-slate-800 font-bold">{endDate}</span>
            </div>
          </div>
        )}
      </div>

      {/* 2. PRINT HEADER (Visible only on print) */}
      <div className="hidden print:block mb-4 p-4 border-b-2 border-slate-800 text-right">
        <div className="flex justify-between items-center mb-2">
          <div>
            <h1 className="text-xl font-bold text-slate-900">الرويال للتطوير العقاري - الإدارة المالية</h1>
            <h2 className="text-sm font-semibold text-slate-600">
              تقرير كشف حساب تحليلي تفصيلي (Subsidiary Ledger Statement)
            </h2>
          </div>
          <div className="text-left text-xs font-mono text-slate-500">
            <div>تاريخ الاستخراج: {new Date().toLocaleDateString('ar-EG')}</div>
            <div>نظام الحسابات المالية الموحد</div>
          </div>
        </div>
        {currentAccount && (
          <div className="bg-slate-100 p-3 rounded border border-slate-300 text-xs flex justify-between items-center mt-2">
            <div>
              <span className="font-bold">الحساب: </span>
              <span className="font-mono font-bold">[{currentAccount.code}] </span>
              <span className="font-bold">{currentAccount.name}</span>
              {currentAccount.parentName && <span> (الحساب العام: {currentAccount.parentName})</span>}
            </div>
            <div>
              <span className="font-bold">الفترة: </span>
              <span className="font-mono">من {startDate} إلى {endDate}</span>
              <span className="mr-3 font-bold">العملة: </span>
              <span>{currency === 'EGP' ? 'جنيه سوداني' : currency === 'USD' ? 'دولار أمريكي ($)' : 'الجنيه والدولار معاً'}</span>
            </div>
          </div>
        )}
      </div>

      {/* 4. TRANSACTIONS TABLE */}
      <div className="enterprise-card rounded-2xl overflow-hidden bg-white border border-slate-200">
        {!selectedAccountCode ? (
          <div className="p-12 text-center text-slate-400 space-y-3">
            <Search className="w-10 h-10 mx-auto opacity-40 text-blue-600" />
            <h3 className="text-sm font-bold text-slate-700">يرجى تحديد أو إدخال كود الحساب الفرعي</h3>
            <p className="text-xs text-slate-400 max-w-md mx-auto">
              اكتب رقم الحساب مباشرة في الحقل بالأعلى (مثلاً: <span className="font-mono font-bold text-blue-600">01060103001</span>) أو اختر من القائمة المنسدلة لعرض كشف الحساب التفصيلي.
            </p>
          </div>
        ) : loading ? (
          <div className="p-12 text-center text-slate-500 space-y-3">
            <RefreshCw className="w-8 h-8 animate-spin mx-auto text-blue-600" />
            <p className="text-xs font-semibold">جارٍ استخراج حركات وقيود كشف الحساب من قاعدة البيانات...</p>
          </div>
        ) : !data ? (
          <div className="p-12 text-center text-slate-400">
            لا توجد بيانات متاحة لهذا الحساب
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-right text-xs">
              <thead className="bg-slate-100 text-slate-700 font-bold border-b border-slate-300">
                <tr>
                  <th className="py-2.5 px-3 border-l border-slate-200 whitespace-nowrap w-24">التاريخ</th>
                  <th className="py-2.5 px-3 border-l border-slate-200 whitespace-nowrap w-28">رقم القيد / السند</th>
                  <th className="py-2.5 px-3 border-l border-slate-200 whitespace-nowrap w-20">نوع السند</th>
                  <th className="py-2.5 px-3 border-l border-slate-200 min-w-[240px]">البيان والشرح والتفاصيل</th>

                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <>
                      <th className="py-2.5 px-3 text-left font-mono text-rose-800 bg-rose-50/50 whitespace-nowrap">مدين ج.س</th>
                      <th className="py-2.5 px-3 text-left font-mono text-emerald-800 bg-emerald-50/50 border-l border-slate-200 whitespace-nowrap">دائن ج.س</th>
                    </>
                  )}

                  {(currency === 'USD' || currency === 'BOTH') && (
                    <>
                      <th className="py-2.5 px-3 text-left font-mono text-blue-800 bg-blue-50/50 whitespace-nowrap">مدين $</th>
                      <th className="py-2.5 px-3 text-left font-mono text-blue-800 bg-blue-50/50 border-l border-slate-200 whitespace-nowrap">دائن $</th>
                    </>
                  )}

                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <th className="py-2.5 px-3 text-left font-mono bg-slate-200/60 border-l border-slate-200 whitespace-nowrap">
                      الرصيد التراكمي (ج.س)
                    </th>
                  )}

                  {(currency === 'USD' || currency === 'BOTH') && (
                    <th className="py-2.5 px-3 text-left font-mono bg-blue-100/50 whitespace-nowrap">
                      الرصيد التراكمي ($)
                    </th>
                  )}
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-100 font-mono">
                {/* Row 0: Opening Balance Row */}
                <tr className="bg-slate-50 font-bold border-b border-slate-300">
                  <td className="py-2.5 px-3 text-slate-500 whitespace-nowrap">{startDate}</td>
                  <td className="py-2.5 px-3 text-slate-400 font-sans whitespace-nowrap">-</td>
                  <td className="py-2.5 px-3 text-blue-700 font-sans whitespace-nowrap">
                    <span className="bg-blue-100 text-blue-800 px-2 py-0.5 rounded text-[11px]">رصيد سابق</span>
                  </td>
                  <td className="py-2.5 px-3 font-sans text-slate-700 font-medium">
                    الرصيد الافتتاحي السابق المدور لحساب [{data.account.name}] عن فترة ما قبل {startDate}
                  </td>

                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <>
                      <td className="py-2.5 px-3 text-left text-slate-400 whitespace-nowrap">-</td>
                      <td className="py-2.5 px-3 text-left text-slate-400 border-l border-slate-200 whitespace-nowrap">-</td>
                    </>
                  )}

                  {(currency === 'USD' || currency === 'BOTH') && (
                    <>
                      <td className="py-2.5 px-3 text-left text-slate-400 whitespace-nowrap">-</td>
                      <td className="py-2.5 px-3 text-left text-slate-400 border-l border-slate-200 whitespace-nowrap">-</td>
                    </>
                  )}

                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <td className="py-2.5 px-3 text-left bg-slate-200/50 border-l border-slate-200 whitespace-nowrap">
                      {renderNetBadge(data.netOpen, false, 'ج.س')}
                    </td>
                  )}

                  {(currency === 'USD' || currency === 'BOTH') && (
                    <td className="py-2.5 px-3 text-left bg-blue-100/40 whitespace-nowrap">
                      {renderNetBadge(data.netOpenDolar, false, '$')}
                    </td>
                  )}
                </tr>

                {/* Transactions rows */}
                {data.transactions.length === 0 ? (
                  <tr>
                    <td colSpan={10} className="py-8 text-center text-slate-400 font-sans">
                      لا توجد حركات مسجلة لهذا الحساب خلال الفترة المحددة (الرصيد منقول فقط)
                    </td>
                  </tr>
                ) : (
                  data.transactions.map((t, idx) => (
                    <tr key={`tx-${t.Id}-${idx}`} className="hover:bg-blue-50/40 transition">
                      <td className="py-2 px-3 text-slate-600 whitespace-nowrap">
                        {formatDate(t.TheDate)}
                      </td>
                      <td className="py-2 px-3 whitespace-nowrap">
                        <button
                          type="button"
                          onClick={() => openVoucherModal(t.ReqNum)}
                          className="font-bold text-blue-600 hover:text-blue-800 flex items-center gap-1 cursor-pointer"
                          title="انقر لاستعراض سند القيد كاملاً"
                        >
                          <span>{t.ReqNum}</span>
                          <ExternalLink className="w-3 h-3 opacity-60 no-print" />
                        </button>
                      </td>
                      <td className="py-2 px-3 font-sans whitespace-nowrap">
                        <span className="bg-slate-100 text-slate-700 px-2 py-0.5 rounded text-[11px] font-semibold border border-slate-200">
                          {t.DocType || 'قيد'}
                        </span>
                      </td>
                      <td className="py-2 px-3 font-sans text-slate-700 max-w-md truncate" title={t.Detail}>
                        {t.Detail || '-'}
                      </td>

                      {/* Movements EGP */}
                      {(currency === 'EGP' || currency === 'BOTH') && (
                        <>
                          <td className="py-2 px-3 text-left font-bold text-rose-700 bg-rose-50/20 whitespace-nowrap">
                            {t.Debtor > 0 ? formatCurrency(t.Debtor) : '-'}
                          </td>
                          <td className="py-2 px-3 text-left font-bold text-emerald-700 bg-emerald-50/20 border-l border-slate-200 whitespace-nowrap">
                            {t.Creditor > 0 ? formatCurrency(t.Creditor) : '-'}
                          </td>
                        </>
                      )}

                      {/* Movements USD */}
                      {(currency === 'USD' || currency === 'BOTH') && (
                        <>
                          <td className="py-2 px-3 text-left font-bold text-blue-700 bg-blue-50/20 whitespace-nowrap">
                            {t.DebtorDolar > 0 ? `$${formatCurrency(t.DebtorDolar)}` : '-'}
                          </td>
                          <td className="py-2 px-3 text-left font-bold text-blue-700 bg-blue-50/20 border-l border-slate-200 whitespace-nowrap">
                            {t.CreditorDolar > 0 ? `$${formatCurrency(t.CreditorDolar)}` : '-'}
                          </td>
                        </>
                      )}

                      {/* Running Net Balances */}
                      {(currency === 'EGP' || currency === 'BOTH') && (
                        <td className="py-2 px-3 text-left bg-slate-100/50 border-l border-slate-200 whitespace-nowrap">
                          {renderNetBadge(t.runningNet, false, 'ج.س')}
                        </td>
                      )}
                      {(currency === 'USD' || currency === 'BOTH') && (
                        <td className="py-2 px-3 text-left bg-blue-50/40 whitespace-nowrap">
                          {renderNetBadge(t.runningNetDolar, false, '$')}
                        </td>
                      )}
                    </tr>
                  ))
                )}
              </tbody>

              {/* Table Subtotal Footer */}
              <tfoot className="bg-slate-200/90 font-bold border-t-2 border-slate-400 text-xs">
                <tr>
                  <td colSpan={4} className="py-3 px-3 font-sans text-slate-900 text-sm">
                    إجمالي حركات الفترة وصافي رصيد الحساب [{data.account.name}]:
                  </td>

                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <>
                      <td className="py-3 px-3 text-left text-rose-800 bg-rose-100/50 whitespace-nowrap">
                        {formatCurrency(data.periodDebtor)}
                      </td>
                      <td className="py-3 px-3 text-left text-emerald-800 bg-emerald-100/50 border-l border-slate-300 whitespace-nowrap">
                        {formatCurrency(data.periodCreditor)}
                      </td>
                    </>
                  )}

                  {(currency === 'USD' || currency === 'BOTH') && (
                    <>
                      <td className="py-3 px-3 text-left text-blue-800 bg-blue-100/50 whitespace-nowrap">
                        ${formatCurrency(data.periodDebtorDolar)}
                      </td>
                      <td className="py-3 px-3 text-left text-blue-800 bg-blue-100/50 border-l border-slate-300 whitespace-nowrap">
                        ${formatCurrency(data.periodCreditorDolar)}
                      </td>
                    </>
                  )}

                  {(currency === 'EGP' || currency === 'BOTH') && (
                    <td className="py-3 px-3 text-left bg-slate-300/50 border-l border-slate-300 whitespace-nowrap">
                      {renderNetBadge(data.netEnding, true, 'ج.س')}
                    </td>
                  )}

                  {(currency === 'USD' || currency === 'BOTH') && (
                    <td className="py-3 px-3 text-left bg-blue-100/60 whitespace-nowrap">
                      {renderNetBadge(data.netEndingDolar, true, '$')}
                    </td>
                  )}
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </div>

      {/* 5. PRINT FOOTER (Visible only on print) */}
      <div className="hidden print:block mt-8 pt-4 border-t border-slate-400 text-xs">
        <div className="grid grid-cols-3 gap-8 text-center">
          <div>
            <p className="font-bold text-slate-800">إعداد / المحاسب المسؤول</p>
            <p className="mt-8 text-slate-400">التوقيع: ........................................</p>
          </div>
          <div>
            <p className="font-bold text-slate-800">مراجعة / المراجع المالي</p>
            <p className="mt-8 text-slate-400">التوقيع: ........................................</p>
          </div>
          <div>
            <p className="font-bold text-slate-800">اعتماد / المدير المالي (CFO)</p>
            <p className="mt-8 text-slate-400">التوقيع: ........................................</p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default SubAccountLedgerView;

