import React, { useState, useEffect, useMemo } from 'react';
import { useLocation, useNavigate, Routes, Route, Navigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { exportStyledTrialBalanceExcel, exportStyledTransactionsExcel, exportStyledGeneralLedgerExcel, exportStyledSubAccountLedgerExcel } from './utils/excelExport';
import GeneralLedgerView from './components/GeneralLedgerView';
import SubAccountLedgerView, { SubAccountItem, SubAccountLedgerData } from './components/SubAccountLedgerView';
import LoginView, { UserSession } from './components/LoginView';
import { 
  LayoutDashboard, 
  FolderTree, 
  ReceiptText, 
  Search, 
  RefreshCw, 
  TrendingUp, 
  TrendingDown, 
  Layers, 
  ChevronRight, 
  ChevronLeft, 
  ChevronDown,
  Download, 
  Printer, 
  ShieldCheck, 
  CheckCircle2, 
  Clock, 
  ArrowDownLeft, 
  X, 
  FileSpreadsheet, 
  Building2, 
  Scale,
  DollarSign,
  ChevronUp,
  Folder,
  FolderOpen,
  FileText,
  Tag,
  ListTree,
  FilePieChart,
  Calendar,
  SlidersHorizontal,
  Coins,
  BookOpen,
  CheckSquare,
  Square,
  Filter,
  LogOut,
  User as UserIcon
} from 'lucide-react';

interface Stats {
  totalAccounts: number;
  totalTransactions: number;
  totalEntries: number;
  totalDebtor: number;
  totalCreditor: number;
  totalDebtorDolar: number;
  totalCreditorDolar: number;
}

interface RecentTx {
  Id: number;
  ReqNum: string;
  DocType: string;
  TheDate: string;
  AccountName: string;
  Debtor: number;
  Creditor: number;
  DebtorDolar: number;
  CreditorDolar: number;
  Curruncy_Name: string;
  Detail: string;
}

interface DocTypeStat {
  DocType: string;
  count: number;
  totalDebtor: number;
  totalCreditor: number;
  totalDebtorDolar: number;
  totalCreditorDolar: number;
}

interface GroupedEntry {
  ReqNum: string;
  itemsCount: number;
  docType: string;
  theDate: string;
  totalDebtor: number;
  totalCreditor: number;
  totalDebtorDolar: number;
  totalCreditorDolar: number;
  sampleDetail: string;
  curruncyName: string;
}

interface EntryLine {
  Id: number;
  ReqNum: string;
  DocType: string;
  TheDate: string;
  AccountNumber: string;
  AccountName: string;
  PAccountName: string;
  Debtor: number;
  Creditor: number;
  DebtorDolar: number;
  CreditorDolar: number;
  Current_blance: number;
  Current_Dolar: number;
  Curruncy_Name: string;
  Check_Number: string;
  Detail: string;
}

interface RawAccount {
  Id: number;
  AccountType: string;
  SubAccountName: string;
  SubAccountNum: string;
  AccountName: string;
  AccountNum: string;
}

interface TreeNode {
  id: number;
  code: string;
  name: string;
  parentCode: string;
  parentName: string;
  type: string;
  isGeneral: boolean;
  children: TreeNode[];
  level: number;
}

interface TrialBalanceRow {
  AccountNumber: string;
  AccountName: string;
  ParentGroupName: string;
  ParentGroupNum: string;
  SubParentName?: string;
  SubParentNum?: string;
  mainParentNum?: string;
  mainParentName?: string;
  AccountType: string;
  category: string;
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

interface TrialBalanceGroup {
  groupKey: string;
  groupCode: string;
  groupName: string;
  category: string;
  rows: TrialBalanceRow[];
  totals: {
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
  };
}

interface GeneralAccountItem {
  code: string;
  name: string;
  parentCode: string;
  parentName: string;
  category: string;
}

interface GeneralLedgerTransaction {
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

interface GeneralLedgerChildAccount {
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

interface GeneralLedgerAccountData {
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
  childAccountsCount?: number;
  childAccounts?: GeneralLedgerChildAccount[];
}

interface GeneralLedgerResponse {
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

export default function App() {
  const location = useLocation();
  const navigate = useNavigate();

  // Authentication State
  const [currentUser, setCurrentUser] = useState<UserSession | null>(() => {
    try {
      const saved = localStorage.getItem('alroyal_session') || sessionStorage.getItem('alroyal_session');
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  });

  const handleLogout = () => {
    localStorage.removeItem('alroyal_session');
    sessionStorage.removeItem('alroyal_session');
    setCurrentUser(null);
  };

  // Active tab derived from current URL pathname
  const activeTab = useMemo(() => {
    const path = location.pathname.replace(/^\//, '');
    if (!path || path === 'dashboard') return 'dashboard';
    if (path.startsWith('transactions')) return 'transactions';
    if (path.startsWith('accounts')) return 'accounts';
    if (path.startsWith('reports')) return 'reports';
    return 'dashboard';
  }, [location.pathname]);
  
  // Dashboard state
  const [stats, setStats] = useState<Stats | null>(null);
  const [recentTx, setRecentTx] = useState<RecentTx[]>([]);
  const [docTypes, setDocTypes] = useState<DocTypeStat[]>([]);
  const [loadingStats, setLoadingStats] = useState(true);

  // Accounts state
  const [rawAccounts, setRawAccounts] = useState<RawAccount[]>([]);
  const [accountSearch, setAccountSearch] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<'all' | '01' | '02' | '03' | '04'>('all');
  const [selectedAccountTypeFilter, setSelectedAccountTypeFilter] = useState<'all' | 'general' | 'detail'>('all');
  const [loadingAccounts, setLoadingAccounts] = useState(false);
  const [expandedNodes, setExpandedNodes] = useState<{ [code: string]: boolean }>({});
  const [selectedAccountNode, setSelectedAccountNode] = useState<TreeNode | null>(null);

  // Grouped Transactions state
  const [groupedEntries, setGroupedEntries] = useState<GroupedEntry[]>([]);
  const [entryPage, setEntryPage] = useState(1);
  const [entryTotalPages, setEntryTotalPages] = useState(1);
  const [entryTotal, setEntryTotal] = useState(0);
  const [entrySearch, setEntrySearch] = useState('');
  const [entryDocType, setEntryDocType] = useState('');
  const [loadingEntries, setLoadingEntries] = useState(false);

  // Expanded Entries state & cached details
  const [expandedReqNums, setExpandedReqNums] = useState<{ [key: string]: boolean }>({});
  const [entryDetailsCache, setEntryDetailsCache] = useState<{ [key: string]: EntryLine[] }>({});
  const [loadingReqDetails, setLoadingReqDetails] = useState<{ [key: string]: boolean }>({});

  // Full Voucher Modal
  const [modalEntry, setModalEntry] = useState<{ reqNum: string; lines: EntryLine[] } | null>(null);

  // REPORT: TRIAL BALANCE STATE
  const [tbStartDate, setTbStartDate] = useState('2021-01-01');
  const [tbEndDate, setTbEndDate] = useState('2023-12-31');
  const [tbCurrency, setTbCurrency] = useState<'EGP' | 'USD' | 'BOTH'>('EGP');
  const [tbMode, setTbMode] = useState<'detailed' | 'summary'>('detailed');
  const [tbGroupBy, setTbGroupBy] = useState<'mainParent' | 'root' | 'subParent' | 'none'>('mainParent');
  const [collapsedGroups, setCollapsedGroups] = useState<{ [key: string]: boolean }>({});
  const [tbSearch, setTbSearch] = useState('');
  const [tbCategory, setTbCategory] = useState<string>('all');
  const [tbRows, setTbRows] = useState<TrialBalanceRow[]>([]);
  const [tbLoading, setTbLoading] = useState(false);

  // REPORT SUB-TAB
  const [reportTab, setReportTab] = useState<'trialBalance' | 'generalLedger' | 'subAccountLedger'>('generalLedger');

  // REPORT: GENERAL LEDGER STATE
  const [glStartDate, setGlStartDate] = useState('2023-01-01');
  const [glEndDate, setGlEndDate] = useState('2023-12-31');
  const [glCurrency, setGlCurrency] = useState<'EGP' | 'USD' | 'BOTH'>('EGP');
  const [glMode, setGlMode] = useState<'detailed' | 'summary'>('detailed');
  const [glGeneralAccounts, setGlGeneralAccounts] = useState<GeneralAccountItem[]>([]);
  const [glSelectedCodes, setGlSelectedCodes] = useState<string[]>([]);
  const [glAccountSearch, setGlAccountSearch] = useState('');
  const [glAccountPickerOpen, setGlAccountPickerOpen] = useState(false);
  const [glSelectedCategory, setGlSelectedCategory] = useState<string>('all');
  const [glData, setGlData] = useState<GeneralLedgerResponse | null>(null);
  const [glLoading, setGlLoading] = useState(false);
  const [glCollapsedAccounts, setGlCollapsedAccounts] = useState<{ [code: string]: boolean }>({});

  // REPORT: SUB-ACCOUNT LEDGER STATE
  const [subAccountsList, setSubAccountsList] = useState<SubAccountItem[]>([]);
  const [selectedSubAccountCode, setSelectedSubAccountCode] = useState<string>('');
  const [subStartDate, setSubStartDate] = useState<string>('2021-01-01');
  const [subEndDate, setSubEndDate] = useState<string>('2023-12-31');
  const [subCurrency, setSubCurrency] = useState<'EGP' | 'USD' | 'BOTH'>('EGP');
  const [subLedgerData, setSubLedgerData] = useState<SubAccountLedgerData | null>(null);
  const [subLoading, setSubLoading] = useState(false);

  // Live Time
  const [currentTime, setCurrentTime] = useState(new Date().toLocaleTimeString('ar-SA'));

  useEffect(() => {
    const timer = setInterval(() => {
      setCurrentTime(new Date().toLocaleTimeString('ar-SA'));
    }, 1000);
    return () => clearInterval(timer);
  }, []);

  // Fetch Dashboard Stats
  const loadDashboard = async () => {
    setLoadingStats(true);
    try {
      const res = await fetch('/api/dashboard/stats');
      const data = await res.json();
      setStats(data.stats);
      setRecentTx(data.recentTransactions || []);
      setDocTypes(data.docTypes || []);
    } catch (err) {
      console.error(err);
    } finally {
      setLoadingStats(false);
    }
  };

  // Fetch All Accounts
  const loadAccounts = async () => {
    setLoadingAccounts(true);
    try {
      const res = await fetch('/api/accounts');
      const data = await res.json();
      const accountsList = data.accounts || [];
      setRawAccounts(accountsList);
      
      // Default: Collapse all accounts tree branches
      setExpandedNodes({});
    } catch (err) {
      console.error(err);
    } finally {
      setLoadingAccounts(false);
    }
  };

  // Fetch Grouped Entries
  const loadGroupedEntries = async () => {
    setLoadingEntries(true);
    try {
      const query = new URLSearchParams({
        page: entryPage.toString(),
        limit: '20'
      });
      if (entrySearch) query.append('search', entrySearch);
      if (entryDocType) query.append('docType', entryDocType);
      const res = await fetch(`/api/transactions/grouped?${query.toString()}`);
      const data = await res.json();
      setGroupedEntries(data.data || []);
      setEntryTotalPages(data.totalPages || 1);
      setEntryTotal(data.total || 0);
    } catch (err) {
      console.error(err);
    } finally {
      setLoadingEntries(false);
    }
  };

  // Fetch Trial Balance Report
  const loadTrialBalance = async () => {
    setTbLoading(true);
    try {
      const query = new URLSearchParams({
        startDate: tbStartDate,
        endDate: tbEndDate
      });
      const res = await fetch(`/api/reports/trial-balance?${query.toString()}`);
      const data = await res.json();
      setTbRows(data.data || []);
    } catch (err) {
      console.error(err);
    } finally {
      setTbLoading(false);
    }
  };

  // Fetch General Accounts List for Multi-select
  const loadGeneralAccounts = async () => {
    try {
      const res = await fetch('/api/accounts/general');
      const data = await res.json();
      setGlGeneralAccounts(data.accounts || []);
    } catch (err) {
      console.error(err);
    }
  };

  // Fetch General Ledger Report Data
  const loadGeneralLedger = async () => {
    if (!glSelectedCodes.length) {
      setGlData(null);
      return;
    }
    setGlLoading(true);
    try {
      const query = new URLSearchParams({
        startDate: glStartDate,
        endDate: glEndDate,
        accountCodes: glSelectedCodes.join(','),
        mode: glMode
      });
      const res = await fetch(`/api/reports/general-ledger?${query.toString()}`);
      const data = await res.json();
      setGlData(data);
    } catch (err) {
      console.error(err);
    } finally {
      setGlLoading(false);
    }
  };

  const toggleGlAccount = (code: string) => {
    setGlSelectedCodes(prev => {
      if (prev.includes(code)) {
        return prev.filter(c => c !== code);
      } else {
        return [...prev, code];
      }
    });
  };

  const selectAllGlAccounts = () => {
    const codes = filteredGeneralAccounts.map(a => a.code);
    setGlSelectedCodes(prev => Array.from(new Set([...prev, ...codes])));
  };

  const clearAllGlAccounts = () => {
    setGlSelectedCodes([]);
  };

  const toggleCollapseGlAccount = (code: string) => {
    setGlCollapsedAccounts(prev => ({ ...prev, [code]: !prev[code] }));
  };

  // Export General Ledger to Styled Excel
  const exportGeneralLedgerExcel = async () => {
    if (!glData || !glData.accounts.length) return;
    try {
      await exportStyledGeneralLedgerExcel({
        accounts: glData.accounts,
        grandTotals: glData.grandTotals,
        startDate: glStartDate,
        endDate: glEndDate,
        currency: glCurrency,
        mode: glMode
      });
    } catch (err) {
      console.error('Error exporting General Ledger Excel:', err);
    }
  };

  // Load Sub-Accounts List
  const loadSubAccountsList = async () => {
    try {
      const res = await fetch('/api/accounts/sub-accounts');
      const data = await res.json();
      setSubAccountsList(data.accounts || []);
    } catch (err) {
      console.error('Error fetching sub-accounts:', err);
    }
  };

  // Load Sub-Account Ledger
  const loadSubAccountLedger = async (codeToLoad?: string) => {
    const code = codeToLoad || selectedSubAccountCode;
    if (!code) return;
    setSubLoading(true);
    try {
      const params = new URLSearchParams({
        accountCode: code,
        startDate: subStartDate,
        endDate: subEndDate
      });
      const res = await fetch(`/api/reports/sub-account-ledger?${params.toString()}`);
      const data = await res.json();
      setSubLedgerData(data);
    } catch (err) {
      console.error('Error fetching sub-account ledger:', err);
    } finally {
      setSubLoading(false);
    }
  };

  // Export Sub-Account Ledger to Styled Excel
  const exportSubAccountExcel = async () => {
    if (!subLedgerData) return;
    try {
      await exportStyledSubAccountLedgerExcel({
        account: subLedgerData.account,
        openDebtor: subLedgerData.openDebtor,
        openCreditor: subLedgerData.openCreditor,
        openDebtorDolar: subLedgerData.openDebtorDolar,
        openCreditorDolar: subLedgerData.openCreditorDolar,
        netOpen: subLedgerData.netOpen,
        netOpenDolar: subLedgerData.netOpenDolar,
        periodDebtor: subLedgerData.periodDebtor,
        periodCreditor: subLedgerData.periodCreditor,
        periodDebtorDolar: subLedgerData.periodDebtorDolar,
        periodCreditorDolar: subLedgerData.periodCreditorDolar,
        totalDebtor: subLedgerData.totalDebtor,
        totalCreditor: subLedgerData.totalCreditor,
        netEnding: subLedgerData.netEnding,
        netEndingDolar: subLedgerData.netEndingDolar,
        transactions: subLedgerData.transactions,
        startDate: subStartDate,
        endDate: subEndDate,
        currency: subCurrency
      });
    } catch (err) {
      console.error('Error exporting Sub-Account Ledger Excel:', err);
    }
  };

  // Toggle row expansion in Transactions
  const toggleExpandEntry = async (reqNum: string) => {
    const isCurrentlyExpanded = !!expandedReqNums[reqNum];
    setExpandedReqNums(prev => ({ ...prev, [reqNum]: !isCurrentlyExpanded }));

    if (!isCurrentlyExpanded && !entryDetailsCache[reqNum]) {
      setLoadingReqDetails(prev => ({ ...prev, [reqNum]: true }));
      try {
        const res = await fetch(`/api/transactions/entry/${encodeURIComponent(reqNum)}`);
        const lines: EntryLine[] = await res.json();
        setEntryDetailsCache(prev => ({ ...prev, [reqNum]: lines }));
      } catch (err) {
        console.error(err);
      } finally {
        setLoadingReqDetails(prev => ({ ...prev, [reqNum]: false }));
      }
    }
  };

  // Open Full Voucher Modal
  const openVoucherModal = async (reqNum: string) => {
    let lines = entryDetailsCache[reqNum];
    if (!lines) {
      try {
        const res = await fetch(`/api/transactions/entry/${encodeURIComponent(reqNum)}`);
        lines = await res.json();
        setEntryDetailsCache(prev => ({ ...prev, [reqNum]: lines }));
      } catch (err) {
        console.error(err);
        return;
      }
    }
    setModalEntry({ reqNum, lines });
  };

  useEffect(() => {
    loadDashboard();
  }, []);

  useEffect(() => {
    if (activeTab === 'accounts' && rawAccounts.length === 0) loadAccounts();
  }, [activeTab]);

  useEffect(() => {
    if (activeTab === 'transactions') loadGroupedEntries();
  }, [activeTab, entryPage, entryDocType]);

  useEffect(() => {
    if (activeTab === 'reports') {
      if (glGeneralAccounts.length === 0) loadGeneralAccounts();
      if (reportTab === 'trialBalance' && tbRows.length === 0) loadTrialBalance();
      if (reportTab === 'generalLedger') loadGeneralLedger();
      if (reportTab === 'subAccountLedger') {
        if (subAccountsList.length === 0) loadSubAccountsList();
        if (selectedSubAccountCode) loadSubAccountLedger();
      }
    }
  }, [activeTab, reportTab, glStartDate, glEndDate, glMode, glSelectedCodes, selectedSubAccountCode, subStartDate, subEndDate]);

  const filteredGeneralAccounts = useMemo(() => {
    let list = glGeneralAccounts;
    if (glSelectedCategory !== 'all') {
      list = list.filter(a => a.category === glSelectedCategory || a.code.startsWith(glSelectedCategory));
    }
    if (glAccountSearch.trim()) {
      const q = glAccountSearch.toLowerCase().trim();
      list = list.filter(a => a.name.toLowerCase().includes(q) || a.code.includes(q) || a.parentName?.toLowerCase().includes(q));
    }
    return list;
  }, [glGeneralAccounts, glSelectedCategory, glAccountSearch]);

  const formatCurrency = (val: number | undefined) => {
    if (val === undefined || val === null) return '0.00';
    return Number(val).toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  };

  const formatDate = (dateStr: string) => {
    if (!dateStr) return '-';
    try {
      return new Date(dateStr).toISOString().split('T')[0];
    } catch {
      return dateStr;
    }
  };

  const balanceDiff = useMemo(() => {
    if (!stats) return 0;
    return (stats.totalDebtor || 0) - (stats.totalCreditor || 0);
  }, [stats]);

  const balanceDolarDiff = useMemo(() => {
    if (!stats) return 0;
    return (stats.totalDebtorDolar || 0) - (stats.totalCreditorDolar || 0);
  }, [stats]);

  // BUILD HIERARCHICAL TREE FROM RAW ACCOUNTS
  const accountTreeData = useMemo(() => {
    if (!rawAccounts.length) return [];

    const validAccounts = rawAccounts.filter(a => a.SubAccountNum && a.SubAccountNum !== '00');

    const nodeMap = new Map<string, TreeNode>();
    validAccounts.forEach(acc => {
      const isGeneral = acc.AccountType?.includes('عـ') || acc.AccountType === 'عام';
      nodeMap.set(acc.SubAccountNum, {
        id: acc.Id,
        code: acc.SubAccountNum,
        name: acc.SubAccountName || acc.AccountName,
        parentCode: acc.AccountNum,
        parentName: acc.AccountName,
        type: acc.AccountType,
        isGeneral,
        children: [],
        level: 1
      });
    });

    validAccounts.forEach(acc => {
      if (acc.Id === 669) {
        const item = nodeMap.get(acc.SubAccountNum);
        if (item) item.parentCode = '01060401';
      }
    });

    const rootNodes: TreeNode[] = [];

    nodeMap.forEach((node) => {
      const parent = nodeMap.get(node.parentCode);
      if (parent && parent.code !== node.code) {
        node.level = parent.level + 1;
        parent.children.push(node);
      } else if (node.parentCode === '0' || node.parentCode === '00' || ['01', '02', '03', '04'].includes(node.code)) {
        node.level = 1;
        rootNodes.push(node);
      }
    });

    const sortTree = (nodes: TreeNode[]) => {
      nodes.sort((a, b) => a.code.localeCompare(b.code, undefined, { numeric: true }));
      nodes.forEach(n => {
        if (n.children.length) sortTree(n.children);
      });
    };
    sortTree(rootNodes);

    return rootNodes;
  }, [rawAccounts]);

  const expandAllTree = () => {
    const allExpanded: { [code: string]: boolean } = {};
    rawAccounts.forEach(acc => {
      allExpanded[acc.SubAccountNum] = true;
    });
    setExpandedNodes(allExpanded);
  };

  const collapseAllTree = () => {
    setExpandedNodes({});
  };

  const toggleTreeNode = (code: string) => {
    setExpandedNodes(prev => ({ ...prev, [code]: !prev[code] }));
  };

  const filteredTreeData = useMemo(() => {
    if (!accountTreeData.length) return [];

    let roots = accountTreeData;
    if (selectedCategory !== 'all') {
      roots = roots.filter(r => r.code === selectedCategory);
    }

    if (!accountSearch.trim() && selectedAccountTypeFilter === 'all') {
      return roots;
    }

    const term = accountSearch.toLowerCase().trim();

    const filterNode = (node: TreeNode): TreeNode | null => {
      const matchSearch = !term || node.name.toLowerCase().includes(term) || node.code.includes(term);
      const matchType = selectedAccountTypeFilter === 'all' 
        ? true 
        : selectedAccountTypeFilter === 'general' ? node.isGeneral : !node.isGeneral;

      const matchingChildren = node.children
        .map(child => filterNode(child))
        .filter((c): c is TreeNode => c !== null);

      if ((matchSearch && matchType) || matchingChildren.length > 0) {
        return {
          ...node,
          children: matchingChildren
        };
      }
      return null;
    };

    return roots
      .map(r => filterNode(r))
      .filter((r): r is TreeNode => r !== null);
  }, [accountTreeData, selectedCategory, accountSearch, selectedAccountTypeFilter]);

  const countSubTreeLeaves = (node: TreeNode): number => {
    if (!node.children || node.children.length === 0) return 1;
    return node.children.reduce((acc, c) => acc + countSubTreeLeaves(c), 0);
  };

  // Recursive Tree Node Renderer
  const renderTreeNode = (node: TreeNode) => {
    const hasChildren = node.children && node.children.length > 0;
    const isExpanded = !!expandedNodes[node.code] || !!accountSearch.trim();
    const isSelected = selectedAccountNode?.code === node.code;
    const leavesCount = hasChildren ? countSubTreeLeaves(node) : 0;

    return (
      <div key={node.code} className="relative select-none text-xs">
        <div 
          onClick={() => setSelectedAccountNode(node)}
          className={`flex items-center justify-between py-2 px-3 my-1 rounded-xl transition-all cursor-pointer ${
            isSelected 
              ? 'bg-blue-600 text-white font-bold shadow-sm' 
              : 'hover:bg-slate-100/90 text-slate-800'
          }`}
          style={{ paddingRight: `${Math.max((node.level - 1) * 24 + 12, 12)}px` }}
        >
          <div className="flex items-center gap-2.5 flex-1 overflow-hidden">
            {hasChildren ? (
              <button
                onClick={(e) => {
                  e.stopPropagation();
                  toggleTreeNode(node.code);
                }}
                className={`w-5 h-5 rounded flex items-center justify-center transition ${
                  isSelected ? 'text-white hover:bg-blue-700' : 'text-slate-500 hover:bg-slate-200'
                }`}
              >
                {isExpanded ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronLeft className="w-3.5 h-3.5" />}
              </button>
            ) : (
              <span className="w-5 flex items-center justify-center text-slate-300">•</span>
            )}

            {hasChildren ? (
              isExpanded ? (
                <FolderOpen className={`w-4 h-4 flex-shrink-0 ${isSelected ? 'text-white' : 'text-blue-600'}`} />
              ) : (
                <Folder className={`w-4 h-4 flex-shrink-0 ${isSelected ? 'text-white' : 'text-blue-500'}`} />
              )
            ) : (
              <FileText className={`w-4 h-4 flex-shrink-0 ${isSelected ? 'text-white' : 'text-slate-400'}`} />
            )}

            <span className={`font-mono font-bold text-xs ${isSelected ? 'text-blue-100' : 'text-blue-700'}`}>
              {node.code}
            </span>
            <span className="font-sans font-semibold truncate text-sm">
              {node.name}
            </span>
          </div>

          <div className="flex items-center gap-2 flex-shrink-0">
            {hasChildren && (
              <span className={`px-2 py-0.5 rounded-full font-mono text-[10px] font-bold ${
                isSelected ? 'bg-blue-700 text-white' : 'bg-slate-100 text-slate-600 border border-slate-200'
              }`}>
                {leavesCount} حساب فرعي
              </span>
            )}

            <span className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
              node.isGeneral 
                ? isSelected ? 'bg-blue-800 text-white' : 'bg-blue-50 text-blue-700 border border-blue-200' 
                : isSelected ? 'bg-blue-900 text-white' : 'bg-slate-100 text-slate-600 border border-slate-200'
            }`}>
              {node.isGeneral ? 'رئيسي (عام)' : 'تحليلي (تفاصيل)'}
            </span>
          </div>
        </div>

        {hasChildren && isExpanded && (
          <div className="relative border-r border-slate-200/90 mr-4">
            {node.children.map(child => renderTreeNode(child))}
          </div>
        )}
      </div>
    );
  };

  // TRIAL BALANCE COMPUTED ROWS (Summary vs Detailed)
  const processedTrialBalanceRows = useMemo(() => {
    let rows = tbRows;

    // Filter by Category
    if (tbCategory !== 'all') {
      rows = rows.filter(r => r.category === tbCategory);
    }

    // Filter by search
    if (tbSearch.trim()) {
      const q = tbSearch.toLowerCase().trim();
      rows = rows.filter(r => 
        r.AccountName?.toLowerCase().includes(q) || 
        r.AccountNumber?.includes(q) || 
        r.ParentGroupName?.toLowerCase().includes(q)
      );
    }

    if (tbMode === 'detailed') {
      return rows;
    }

    // SUMMARY AGGREGATION BY ParentGroupName
    const groupsMap = new Map<string, TrialBalanceRow>();
    rows.forEach(r => {
      const gKey = r.ParentGroupName || 'عام';
      const existing = groupsMap.get(gKey);
      if (!existing) {
        groupsMap.set(gKey, {
          AccountNumber: r.ParentGroupNum || r.AccountNumber.substring(0, 6),
          AccountName: gKey,
          ParentGroupName: r.category,
          ParentGroupNum: r.ParentGroupNum,
          AccountType: 'مجموعة رئيسية',
          category: r.category,
          openDebtor: r.openDebtor,
          openCreditor: r.openCreditor,
          openDebtorDolar: r.openDebtorDolar,
          openCreditorDolar: r.openCreditorDolar,
          periodDebtor: r.periodDebtor,
          periodCreditor: r.periodCreditor,
          periodDebtorDolar: r.periodDebtorDolar,
          periodCreditorDolar: r.periodCreditorDolar,
          totalDebtor: r.totalDebtor,
          totalCreditor: r.totalCreditor,
          totalDebtorDolar: r.totalDebtorDolar,
          totalCreditorDolar: r.totalCreditorDolar,
          netOpen: r.netOpen,
          netOpenDolar: r.netOpenDolar,
          netEnding: r.netEnding,
          netEndingDolar: r.netEndingDolar
        });
      } else {
        existing.openDebtor += r.openDebtor;
        existing.openCreditor += r.openCreditor;
        existing.openDebtorDolar += r.openDebtorDolar;
        existing.openCreditorDolar += r.openCreditorDolar;
        existing.periodDebtor += r.periodDebtor;
        existing.periodCreditor += r.periodCreditor;
        existing.periodDebtorDolar += r.periodDebtorDolar;
        existing.periodCreditorDolar += r.periodCreditorDolar;
        existing.totalDebtor += r.totalDebtor;
        existing.totalCreditor += r.totalCreditor;
        existing.totalDebtorDolar += r.totalDebtorDolar;
        existing.totalCreditorDolar += r.totalCreditorDolar;
        existing.netOpen += r.netOpen;
        existing.netOpenDolar += r.netOpenDolar;
        existing.netEnding += r.netEnding;
        existing.netEndingDolar += r.netEndingDolar;
      }
    });

    return Array.from(groupsMap.values());
  }, [tbRows, tbMode, tbCategory, tbSearch]);

  // TRIAL BALANCE GRAND TOTALS
  const tbGrandTotals = useMemo(() => {
    return processedTrialBalanceRows.reduce((acc, row) => {
      acc.openDebtor += row.openDebtor;
      acc.openCreditor += row.openCreditor;
      acc.openDebtorDolar += row.openDebtorDolar;
      acc.openCreditorDolar += row.openCreditorDolar;

      acc.periodDebtor += row.periodDebtor;
      acc.periodCreditor += row.periodCreditor;
      acc.periodDebtorDolar += row.periodDebtorDolar;
      acc.periodCreditorDolar += row.periodCreditorDolar;

      acc.totalDebtor += row.totalDebtor;
      acc.totalCreditor += row.totalCreditor;
      acc.totalDebtorDolar += row.totalDebtorDolar;
      acc.totalCreditorDolar += row.totalCreditorDolar;

      // Net Openings
      const netO = (row.openDebtor || 0) - (row.openCreditor || 0);
      if (netO > 0) acc.netOpenDebit += netO;
      else if (netO < 0) acc.netOpenCredit += Math.abs(netO);

      const netODolar = (row.openDebtorDolar || 0) - (row.openCreditorDolar || 0);
      if (netODolar > 0) acc.netOpenDebitDolar += netODolar;
      else if (netODolar < 0) acc.netOpenCreditDolar += Math.abs(netODolar);

      // Net Endings
      const netE = ((row.openDebtor || 0) + (row.periodDebtor || 0)) - ((row.openCreditor || 0) + (row.periodCreditor || 0));
      if (netE > 0) acc.netEndingDebit += netE;
      else if (netE < 0) acc.netEndingCredit += Math.abs(netE);

      const netEDolar = ((row.openDebtorDolar || 0) + (row.periodDebtorDolar || 0)) - ((row.openCreditorDolar || 0) + (row.periodCreditorDolar || 0));
      if (netEDolar > 0) acc.netEndingDebitDolar += netEDolar;
      else if (netEDolar < 0) acc.netEndingCreditDolar += Math.abs(netEDolar);

      return acc;
    }, {
      openDebtor: 0,
      openCreditor: 0,
      openDebtorDolar: 0,
      openCreditorDolar: 0,
      periodDebtor: 0,
      periodCreditor: 0,
      periodDebtorDolar: 0,
      periodCreditorDolar: 0,
      totalDebtor: 0,
      totalCreditor: 0,
      totalDebtorDolar: 0,
      totalCreditorDolar: 0,

      // Net Balances
      netOpenDebit: 0,
      netOpenCredit: 0,
      netOpenDebitDolar: 0,
      netOpenCreditDolar: 0,
      netEndingDebit: 0,
      netEndingCredit: 0,
      netEndingDebitDolar: 0,
      netEndingCreditDolar: 0
    });
  }, [processedTrialBalanceRows]);

  // MAIN PARENT NAMES DICTIONARY FOR ACCURATE ARABIC LABELS
  const MAIN_PARENT_TITLES: { [prefix: string]: string } = {
    '0101': 'الأصول الثابتة',
    '0102': 'الأصول المتداولة (النقدية والبنوك)',
    '0103': 'المشروعات قيد التنفيذ',
    '0104': 'الاعتمادات المستندية',
    '0105': 'أصول أخرى',
    '0106': 'المدينون (العملاء والعهد والمقاولون)',
    '0107': 'مصاريف التأسيس',
    '0201': 'حقوق المساهمين ورأس المال',
    '0202': 'المخصصات وإهلاك الأصول',
    '0203': 'الدائنون (الموردون والتزامات)',
    '0301': 'مصروفات التشغيل المباشرة',
    '0302': 'المصروفات الإدارية والعمومية',
    '0303': 'مصروفات التمويل',
    '0304': 'مخصصات ومصروفات أخرى',
    '0305': 'تكلفة المبيعات والإنتاج',
    '0401': 'إيرادات النشاط والمبيعات',
    '0402': 'إيرادات أخرى متنوعة'
  };

  // GROUPED DATA BY CHOSEN GROUPING (Default: Main Parent Account "الحساب الرئيسي الكبير")
  const tbGroupedData = useMemo<TrialBalanceGroup[]>(() => {
    if (tbGroupBy === 'none') {
      return [{
        groupKey: 'all',
        groupCode: 'ALL',
        groupName: 'كافة الحسابات المالية (عرض مباشر مسطح)',
        category: 'الكل',
        rows: processedTrialBalanceRows,
        totals: tbGrandTotals
      }];
    }

    const map = new Map<string, TrialBalanceGroup>();

    processedTrialBalanceRows.forEach(row => {
      let code = '';
      let name = '';
      let cat = row.category;

      if (tbGroupBy === 'mainParent') {
        const prefix4 = (row.AccountNumber || '').substring(0, 4);
        code = row.mainParentNum || prefix4 || '0100';
        name = row.mainParentName || MAIN_PARENT_TITLES[prefix4] || row.ParentGroupName || 'حساب رئيسي';
      } else if (tbGroupBy === 'root') {
        code = (row.AccountNumber || '').substring(0, 2);
        name = row.category;
      } else if (tbGroupBy === 'subParent') {
        code = row.ParentGroupNum || (row.AccountNumber || '').substring(0, 6);
        name = row.SubParentName || row.ParentGroupName || 'المجموعة الفرعية';
      }

      const key = `${code}_${name}`;
      let grp = map.get(key);
      if (!grp) {
        grp = {
          groupKey: key,
          groupCode: code,
          groupName: name,
          category: cat,
          rows: [],
          totals: {
            openDebtor: 0,
            openCreditor: 0,
            openDebtorDolar: 0,
            openCreditorDolar: 0,
            periodDebtor: 0,
            periodCreditor: 0,
            periodDebtorDolar: 0,
            periodCreditorDolar: 0,
            totalDebtor: 0,
            totalCreditor: 0,
            totalDebtorDolar: 0,
            totalCreditorDolar: 0,
          }
        };
        map.set(key, grp);
      }

      grp.rows.push(row);
      grp.totals.openDebtor += row.openDebtor;
      grp.totals.openCreditor += row.openCreditor;
      grp.totals.openDebtorDolar += row.openDebtorDolar;
      grp.totals.openCreditorDolar += row.openCreditorDolar;
      grp.totals.periodDebtor += row.periodDebtor;
      grp.totals.periodCreditor += row.periodCreditor;
      grp.totals.periodDebtorDolar += row.periodDebtorDolar;
      grp.totals.periodCreditorDolar += row.periodCreditorDolar;
      grp.totals.totalDebtor += row.totalDebtor;
      grp.totals.totalCreditor += row.totalCreditor;
      grp.totals.totalDebtorDolar += row.totalDebtorDolar;
      grp.totals.totalCreditorDolar += row.totalCreditorDolar;
    });

    return Array.from(map.values()).sort((a, b) => a.groupCode.localeCompare(b.groupCode));
  }, [processedTrialBalanceRows, tbGroupBy, tbGrandTotals]);

  // Group Accordion Helpers (COLLAPSED BY DEFAULT)
  const toggleGroup = (key: string) => {
    setCollapsedGroups(prev => {
      const currentlyCollapsed = prev[key] !== false;
      return { ...prev, [key]: !currentlyCollapsed };
    });
  };

  const expandAllGroups = () => {
    const all: { [k: string]: boolean } = {};
    tbGroupedData.forEach(g => { all[g.groupKey] = false; });
    setCollapsedGroups(all);
  };

  const collapseAllGroups = () => {
    setCollapsedGroups({});
  };

  // PROFESSIONAL EXCEL EXPORT (EXCELJS) WITH ROYAL BRAND IDENTITY & CELL FORMATTING
  const exportTrialBalanceExcel = async () => {
    if (!processedTrialBalanceRows.length) return;
    try {
      await exportStyledTrialBalanceExcel({
        tbGroupedData,
        processedTrialBalanceRows,
        tbGrandTotals,
        tbStartDate,
        tbEndDate,
        tbCurrency,
        tbMode,
        tbGroupBy
      });
    } catch (err) {
      console.error('Failed to export styled Excel:', err);
    }
  };

  const exportCSV = (data: any[], filename: string) => {
    if (!data || !data.length) return;
    exportStyledTransactionsExcel(data, filename);
  };

  // Colored Net Balance Badge (Debit in Red/Rose, Credit in Green/Emerald)
  const renderNetBadge = (val: number, isEnding = false, symbol = '') => {
    if (Math.abs(val) < 0.001) {
      return <span className="text-slate-400 font-mono">-</span>;
    }
    const isDebit = val > 0;
    const absVal = Math.abs(val);

    if (isDebit) {
      return (
        <div className="flex items-center gap-1 font-mono">
          <span className={`text-[10px] font-sans font-bold px-1.5 py-0.5 rounded ${
            isEnding 
              ? 'bg-rose-100 text-rose-800 border border-rose-300 font-extrabold shadow-2xs' 
              : 'bg-rose-50 text-rose-700 border border-rose-200/80'
          }`}>
            مدين
          </span>
          <span className={`${isEnding ? 'font-extrabold text-rose-900' : 'font-bold text-rose-700'}`}>
            {formatCurrency(absVal)} {symbol}
          </span>
        </div>
      );
    }

    return (
      <div className="flex items-center gap-1 font-mono">
        <span className={`text-[10px] font-sans font-bold px-1.5 py-0.5 rounded ${
          isEnding 
            ? 'bg-emerald-100 text-emerald-800 border border-emerald-300 font-extrabold shadow-2xs' 
            : 'bg-emerald-50 text-emerald-700 border border-emerald-200/80'
        }`}>
          دائن
        </span>
        <span className={`${isEnding ? 'font-extrabold text-emerald-900' : 'font-bold text-emerald-700'}`}>
          {formatCurrency(absVal)} {symbol}
        </span>
      </div>
    );
  };

  if (!currentUser) {
    return <LoginView onLoginSuccess={(user) => setCurrentUser(user)} />;
  }

  return (
    <div className="min-h-screen bg-enterprise-pattern text-slate-800 flex flex-col font-['IBM_Plex_Sans_Arabic'] selection:bg-blue-600 selection:text-white relative">
      
      {/* GEOMETRIC WATERMARK ACCENTS */}
      <div className="geometric-shapes-overlay" aria-hidden="true">
        <svg className="absolute -top-12 -left-12 w-96 h-96 opacity-[0.06] text-blue-600" viewBox="0 0 200 200" fill="none" stroke="currentColor" strokeWidth="1.2">
          <polygon points="100,20 170,60 170,140 100,180 30,140 30,60" />
          <line x1="100" y1="20" x2="100" y2="100" />
          <line x1="100" y1="100" x2="170" y2="60" />
          <line x1="100" y1="100" x2="30" y2="60" />
          <line x1="100" y1="100" x2="100" y2="180" />
        </svg>

        <svg className="absolute top-1/3 -right-20 w-[480px] h-[480px] opacity-[0.05] text-indigo-700" viewBox="0 0 300 300" fill="none" stroke="currentColor" strokeWidth="1">
          <polygon points="150,30 240,82 240,186 150,238 60,186 60,82" />
          <polygon points="150,60 215,97 215,171 150,208 85,171 85,97" strokeDasharray="3 3" />
        </svg>

        <svg className="absolute -bottom-16 left-1/4 w-[520px] h-[520px] opacity-[0.04] text-emerald-600" viewBox="0 0 400 400" fill="none" stroke="currentColor" strokeWidth="1">
          <rect x="100" y="100" width="200" height="200" transform="rotate(45 200 200)" />
          <circle cx="200" cy="200" r="80" />
        </svg>
      </div>

      {/* TOP APEX BAR (Header) */}
      <header className="h-14 bg-white/95 backdrop-blur-md border-b border-slate-200/90 px-6 flex items-center justify-between text-xs select-none shadow-[0_1px_3px_rgba(0,0,0,0.03)] z-20">
        <div className="flex items-center gap-6">
          <div 
            onClick={() => navigate('/dashboard')} 
            className="flex items-center gap-3 cursor-pointer"
          >
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-600 to-indigo-700 flex items-center justify-center font-bold text-white shadow-md shadow-blue-600/20">
              <Building2 className="w-4 h-4" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="font-bold text-sm tracking-tight text-slate-900">الرويال للتطوير العقاري</span>
                <span className="bg-blue-50 text-blue-700 border border-blue-200/80 px-1.5 py-0.5 rounded text-[10px] font-mono uppercase tracking-wider font-semibold">Enterprise</span>
              </div>
            </div>
          </div>

          <div className="hidden lg:flex items-center gap-2 border-r border-slate-200 pr-6 mr-3 text-slate-500">
            <span>خادم البيانات:</span>
            <span className="font-mono text-slate-700 bg-slate-100 px-2 py-0.5 rounded border border-slate-200 font-semibold">MSSQL 2014 \ RoyalDb</span>
            <span className="w-2 h-2 rounded-full bg-emerald-500 shadow-sm shadow-emerald-500/50 mr-2 animate-pulse"></span>
            <span className="text-emerald-700 font-medium">متصل ونشط</span>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="hidden md:flex items-center gap-2 bg-slate-50 border border-slate-200 px-3 py-1 rounded-lg text-slate-600 font-mono text-[11px]">
            <Clock className="w-3.5 h-3.5 text-blue-600" />
            <span>{currentTime}</span>
          </div>

          <button
            onClick={() => {
              if (activeTab === 'dashboard') loadDashboard();
              if (activeTab === 'accounts') loadAccounts();
              if (activeTab === 'transactions') loadGroupedEntries();
              if (activeTab === 'reports') loadTrialBalance();
            }}
            className="flex items-center gap-1.5 bg-white hover:bg-slate-50 active:bg-slate-100 text-slate-700 px-3.5 py-1.5 rounded-lg border border-slate-300 shadow-xs transition text-xs font-semibold cursor-pointer"
          >
            <RefreshCw className="w-3.5 h-3.5 text-slate-500" />
            <span>مزامنة فورية</span>
          </button>

          {/* User Profile Pill & Quick Logout */}
          <div className="flex items-center gap-2 border-r border-slate-200 pr-3 mr-1">
            <div className="hidden sm:flex flex-col text-left">
              <span className="text-xs font-bold text-slate-800">{currentUser?.displayName || 'مدير النظام'}</span>
              <span className="text-[10px] text-slate-500 font-medium">{currentUser?.role || 'مسؤول النظام'}</span>
            </div>
            <button
              onClick={handleLogout}
              title="تسجيل الخروج من النظام"
              className="flex items-center gap-1.5 bg-red-50 hover:bg-red-100 active:bg-red-200 text-red-700 border border-red-200/80 px-2.5 py-1.5 rounded-lg text-xs font-semibold transition cursor-pointer shadow-xs"
            >
              <LogOut className="w-3.5 h-3.5" />
              <span className="hidden md:inline">خروج</span>
            </button>
          </div>
        </div>
      </header>

      {/* WORKSPACE LAYOUT */}
      <div className="flex-1 flex overflow-hidden z-10">
        {/* SIDEBAR */}
        <aside className="w-64 bg-white/95 backdrop-blur-md border-l border-slate-200/90 flex flex-col justify-between select-none shadow-[2px_0_8px_rgba(15,23,42,0.02)]">
          <div className="py-2">
            {/* Group 1: General Modules */}
            <div className="px-4 pt-2 pb-1.5 text-[11px] font-bold uppercase tracking-wider text-slate-400">
              القوائم والسجلات
            </div>
            <nav className="px-2 space-y-1 mb-4">
              <button
                onClick={() => navigate('/dashboard')}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all ${
                  activeTab === 'dashboard'
                    ? 'bg-blue-600 text-white font-bold shadow-md shadow-blue-600/25'
                    : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900 font-semibold'
                }`}
              >
                <LayoutDashboard className={`w-4 h-4 shrink-0 ${activeTab === 'dashboard' ? 'text-white' : 'text-slate-400'}`} />
                <span>لوحة المؤشرات والرقابة</span>
              </button>

              <button
                onClick={() => navigate('/transactions')}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all ${
                  activeTab === 'transactions'
                    ? 'bg-blue-600 text-white font-bold shadow-md shadow-blue-600/25'
                    : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900 font-semibold'
                }`}
              >
                <ReceiptText className={`w-4 h-4 shrink-0 ${activeTab === 'transactions' ? 'text-white' : 'text-slate-400'}`} />
                <span>دفتر وسجل القيود المحاسبية</span>
              </button>

              <button
                onClick={() => navigate('/accounts')}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all ${
                  activeTab === 'accounts'
                    ? 'bg-blue-600 text-white font-bold shadow-md shadow-blue-600/25'
                    : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900 font-semibold'
                }`}
              >
                <FolderTree className={`w-4 h-4 shrink-0 ${activeTab === 'accounts' ? 'text-white' : 'text-slate-400'}`} />
                <span>دليل شجرة الحسابات الهرمية</span>
              </button>
            </nav>

            {/* Group 2: Financial Reports */}
            <div className="px-4 pt-2 pb-1.5 text-[11px] font-bold uppercase tracking-wider text-slate-400 border-t border-slate-100">
              التقارير المالية المعتمدة
            </div>
            <nav className="px-2 space-y-1">
              <button
                onClick={() => {
                  setReportTab('generalLedger');
                  navigate('/reports');
                }}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all ${
                  activeTab === 'reports' && reportTab === 'generalLedger'
                    ? 'bg-blue-600 text-white font-bold shadow-md shadow-blue-600/25'
                    : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900 font-semibold'
                }`}
              >
                <BookOpen className={`w-4 h-4 shrink-0 ${activeTab === 'reports' && reportTab === 'generalLedger' ? 'text-white' : 'text-slate-400'}`} />
                <span>كشف الحسابات العامة</span>
              </button>

              <button
                onClick={() => {
                  setReportTab('subAccountLedger');
                  navigate('/reports');
                }}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all ${
                  activeTab === 'reports' && reportTab === 'subAccountLedger'
                    ? 'bg-blue-600 text-white font-bold shadow-md shadow-blue-600/25'
                    : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900 font-semibold'
                }`}
              >
                <FileText className={`w-4 h-4 shrink-0 ${activeTab === 'reports' && reportTab === 'subAccountLedger' ? 'text-white' : 'text-slate-400'}`} />
                <span>كشف الحسابات الفرعية</span>
              </button>

              <button
                onClick={() => {
                  setReportTab('trialBalance');
                  navigate('/reports');
                }}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs transition-all ${
                  activeTab === 'reports' && reportTab === 'trialBalance'
                    ? 'bg-blue-600 text-white font-bold shadow-md shadow-blue-600/25'
                    : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900 font-semibold'
                }`}
              >
                <Scale className={`w-4 h-4 shrink-0 ${activeTab === 'reports' && reportTab === 'trialBalance' ? 'text-white' : 'text-slate-400'}`} />
                <span>ميزان المراجعة الشامل</span>
              </button>
            </nav>
          </div>

          {/* User Profile & Logout Box in Sidebar */}
          <div className="p-3 m-2.5 rounded-2xl bg-slate-50 border border-slate-200/90 shadow-xs">
            <div className="flex items-center gap-2.5 mb-2.5">
              <div className="w-9 h-9 rounded-xl bg-gradient-to-tr from-blue-700 via-indigo-600 to-blue-500 flex items-center justify-center text-white shadow-sm shadow-blue-500/20 shrink-0">
                <UserIcon className="w-4 h-4" />
              </div>
              <div className="overflow-hidden">
                <div className="font-bold text-xs text-slate-900 truncate">
                  {currentUser?.displayName || 'مدير النظام'}
                </div>
                <div className="text-[10px] text-slate-500 font-medium truncate">
                  {currentUser?.role || 'مسؤول النظام (Admin)'}
                </div>
              </div>
            </div>

            <button
              onClick={handleLogout}
              className="w-full flex items-center justify-center gap-2 bg-white hover:bg-red-50 text-slate-600 hover:text-red-700 border border-slate-200/90 hover:border-red-200/80 py-1.5 px-3 rounded-xl text-xs font-semibold transition shadow-xs cursor-pointer"
            >
              <LogOut className="w-3.5 h-3.5 text-red-500" />
              <span>تسجيل الخروج</span>
            </button>
          </div>
        </aside>

        {/* MAIN DISPLAY CANVAS */}
        <main className="flex-1 overflow-y-auto p-6 flex flex-col gap-6">
          <Routes>
            <Route path="/" element={<Navigate to="/dashboard" replace />} />

            {/* TAB: DASHBOARD */}
            <Route path="/dashboard" element={
              <div className="space-y-6">
                <div className="enterprise-card rounded-2xl p-6">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4 mb-5">
                    <div>
                      <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                        <Scale className="w-5 h-5 text-blue-600" />
                        <span>ميزان الحركات المالي العام ثنائي العملة (جنيه ودولار)</span>
                      </h2>
                      <p className="text-xs text-slate-500 mt-0.5">
                        مطابقة إجمالي المبالغ المدينة مع المبالغ الدائنة بالجنيه السوداني والدولار الأمريكي لجميع العمليات
                      </p>
                    </div>
                    <div className="flex items-center gap-3">
                      <div className="bg-slate-50 border border-slate-200/90 px-3.5 py-1.5 rounded-xl text-right">
                        <span className="text-[10px] text-slate-500 block font-semibold">فارق الجنيه:</span>
                        <span className="font-mono font-bold text-xs text-emerald-700">
                          {formatCurrency(balanceDiff)} ج.س
                        </span>
                      </div>
                      <div className="bg-blue-50/70 border border-blue-200/90 px-3.5 py-1.5 rounded-xl text-right">
                        <span className="text-[10px] text-blue-700 block font-semibold">فارق الدولار:</span>
                        <span className="font-mono font-bold text-xs text-blue-900">
                          ${formatCurrency(balanceDolarDiff)}
                        </span>
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                    <div className="bg-white border border-slate-200/80 rounded-xl p-4 shadow-xs">
                      <div className="flex items-center justify-between text-slate-500 text-xs font-semibold">
                        <span>إجمالي القيود والسندات</span>
                        <div className="w-8 h-8 rounded-lg bg-blue-50 text-blue-700 flex items-center justify-center">
                          <ReceiptText className="w-4 h-4" />
                        </div>
                      </div>
                      <div className="mt-3 flex items-baseline justify-between">
                        <span className="font-mono text-2xl font-bold text-slate-900">
                          {loadingStats ? '---' : stats?.totalEntries?.toLocaleString('en-US')}
                        </span>
                        <span className="text-[11px] text-slate-500 font-mono">سند وقيد</span>
                      </div>
                      <div className="mt-2 text-[11px] text-slate-500 flex items-center gap-1 font-medium">
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                        <span>تتضمن 14,228 سطر حركة</span>
                      </div>
                    </div>

                    <div className="bg-white border border-slate-200/80 rounded-xl p-4 shadow-xs">
                      <div className="flex items-center justify-between text-slate-500 text-xs font-semibold">
                        <span>دليل شجرة الحسابات</span>
                        <div className="w-8 h-8 rounded-lg bg-indigo-50 text-indigo-700 flex items-center justify-center">
                          <FolderTree className="w-4 h-4" />
                        </div>
                      </div>
                      <div className="mt-3 flex items-baseline justify-between">
                        <span className="font-mono text-2xl font-bold text-slate-900">
                          {loadingStats ? '---' : stats?.totalAccounts?.toLocaleString('en-US')}
                        </span>
                        <span className="text-[11px] text-indigo-700 font-bold font-mono">حساب مفهرس</span>
                      </div>
                      <div className="mt-2 text-[11px] text-slate-500 font-medium">
                        أصول، خصوم، إيرادات، ومصروفات
                      </div>
                    </div>

                    <div className="bg-white border border-slate-200/80 rounded-xl p-4 shadow-xs">
                      <div className="flex items-center justify-between text-slate-500 text-xs font-semibold">
                        <span>إجمالي المدين (SDG & USD)</span>
                        <div className="w-8 h-8 rounded-lg bg-rose-50 text-rose-700 flex items-center justify-center">
                          <TrendingUp className="w-4 h-4" />
                        </div>
                      </div>
                      <div className="mt-2">
                        <div className="font-mono text-lg font-bold text-rose-700 tabular-nums">
                          {loadingStats ? '---' : formatCurrency(stats?.totalDebtor)} <span className="text-xs font-sans text-slate-500">ج.س</span>
                        </div>
                        <div className="font-mono text-sm font-semibold text-blue-700 tabular-nums mt-0.5">
                          ${loadingStats ? '---' : formatCurrency(stats?.totalDebtorDolar)} <span className="text-[10px] font-sans text-slate-500">دولار</span>
                        </div>
                      </div>
                      <div className="mt-2 text-[11px] text-slate-500 font-medium border-t border-slate-100 pt-1.5">
                        المجموع التراكمي لجانب المدين
                      </div>
                    </div>

                    <div className="bg-white border border-slate-200/80 rounded-xl p-4 shadow-xs">
                      <div className="flex items-center justify-between text-slate-500 text-xs font-semibold">
                        <span>إجمالي الدائن (SDG & USD)</span>
                        <div className="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-700 flex items-center justify-center">
                          <TrendingDown className="w-4 h-4" />
                        </div>
                      </div>
                      <div className="mt-2">
                        <div className="font-mono text-lg font-bold text-emerald-700 tabular-nums">
                          {loadingStats ? '---' : formatCurrency(stats?.totalCreditor)} <span className="text-xs font-sans text-slate-500">ج.س</span>
                        </div>
                        <div className="font-mono text-sm font-semibold text-blue-700 tabular-nums mt-0.5">
                          ${loadingStats ? '---' : formatCurrency(stats?.totalCreditorDolar)} <span className="text-[10px] font-sans text-slate-500">دولار</span>
                        </div>
                      </div>
                      <div className="mt-2 text-[11px] text-slate-500 font-medium border-t border-slate-100 pt-1.5">
                        المجموع التراكمي لجانب الدائن
                      </div>
                    </div>
                  </div>
                </div>

                {/* Lower Section */}
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                  <div className="enterprise-card rounded-2xl p-6">
                    <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-100">
                      <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                        <Layers className="w-4 h-4 text-blue-600" />
                        <span>توزيع السندات والعمليات</span>
                      </h3>
                      <span className="text-[10px] font-mono text-slate-400 uppercase font-bold">Classified</span>
                    </div>

                    <div className="space-y-3">
                      {docTypes.map((item, idx) => {
                        const totalAll = stats?.totalTransactions || 1;
                        const percent = ((item.count / totalAll) * 100).toFixed(1);
                        return (
                          <div key={idx} className="bg-slate-50/90 border border-slate-200/70 p-3.5 rounded-xl">
                            <div className="flex items-center justify-between text-xs mb-1.5">
                              <span className="font-bold text-slate-800">{item.DocType}</span>
                              <span className="font-mono text-slate-500 font-medium">{item.count.toLocaleString('en-US')} حركة ({percent}%)</span>
                            </div>
                            <div className="w-full bg-slate-200 h-2 rounded-full overflow-hidden">
                              <div 
                                className="bg-blue-600 h-full rounded-full transition-all duration-500" 
                                style={{ width: `${percent}%` }}
                              ></div>
                            </div>
                            <div className="flex justify-between items-center mt-2.5 text-[11px] text-slate-500 font-mono">
                              <span>المبالغ:</span>
                              <span className="text-slate-900 font-bold">
                                {formatCurrency(item.totalDebtor)} ج.س (${formatCurrency(item.totalDebtorDolar)})
                              </span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  <div className="lg:col-span-2 enterprise-card rounded-2xl p-6 flex flex-col justify-between">
                    <div>
                      <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-100">
                        <div>
                          <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                            <ReceiptText className="w-4 h-4 text-emerald-600" />
                            <span>أحدث أسطر الحركات المسجلة</span>
                          </h3>
                          <span className="text-[11px] text-slate-500 font-medium">آخر 6 أسطر قيود تم ترحيلها</span>
                        </div>
                        <button
                          onClick={() => navigate('/transactions')}
                          className="text-xs text-blue-600 hover:text-blue-700 font-bold flex items-center gap-1 transition"
                        >
                          <span>فتح سجل القيود المجمعة</span>
                          <ArrowDownLeft className="w-3.5 h-3.5" />
                        </button>
                      </div>

                      <div className="overflow-x-auto">
                        <table className="w-full text-right text-xs">
                          <thead>
                            <tr className="border-b border-slate-200 text-slate-500 font-bold">
                              <th className="pb-3 pr-2 font-mono">رقم السند</th>
                              <th className="pb-3 px-2">النوع</th>
                              <th className="pb-3 px-2 font-mono">التاريخ</th>
                              <th className="pb-3 px-2">الحساب</th>
                              <th className="pb-3 px-2 font-mono text-left">مدين (جنيه / $)</th>
                              <th className="pb-3 px-2 font-mono text-left">دائن (جنيه / $)</th>
                              <th className="pb-3 pl-2">البيان</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-mono">
                            {recentTx.map((tx) => (
                              <tr 
                                key={tx.Id} 
                                onClick={() => openVoucherModal(tx.ReqNum)}
                                className="hover:bg-blue-50/50 cursor-pointer transition select-none"
                              >
                                <td className="py-2.5 pr-2 font-bold text-blue-700">{tx.ReqNum || tx.Id}</td>
                                <td className="py-2.5 px-2 font-sans">
                                  <span className="bg-slate-100 text-slate-700 px-2 py-0.5 rounded text-[10px] border border-slate-200 font-semibold">
                                    {tx.DocType || 'سند'}
                                  </span>
                                </td>
                                <td className="py-2.5 px-2 text-slate-500">{formatDate(tx.TheDate)}</td>
                                <td className="py-2.5 px-2 font-sans font-bold text-slate-800">{tx.AccountName}</td>
                                <td className="py-2.5 px-2 text-rose-700 font-bold text-left">
                                  {tx.Debtor > 0 ? (
                                    <div>
                                      <span>{formatCurrency(tx.Debtor)} ج.س</span>
                                      {tx.DebtorDolar > 0 && <span className="block text-[10px] text-blue-700 font-normal">${formatCurrency(tx.DebtorDolar)}</span>}
                                    </div>
                                  ) : '-'}
                                </td>
                                <td className="py-2.5 px-2 text-emerald-700 font-bold text-left">
                                  {tx.Creditor > 0 ? (
                                    <div>
                                      <span>{formatCurrency(tx.Creditor)} ج.س</span>
                                      {tx.CreditorDolar > 0 && <span className="block text-[10px] text-blue-700 font-normal">${formatCurrency(tx.CreditorDolar)}</span>}
                                    </div>
                                  ) : '-'}
                                </td>
                                <td className="py-2.5 pl-2 font-sans text-slate-500 max-w-[160px] truncate" title={tx.Detail}>
                                  {tx.Detail || '-'}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500">
                      <span className="flex items-center gap-1.5 font-medium">
                        <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                        <span>انقر على أي سطر لعرض قسيمة القيد المحاسبي الكاملة</span>
                      </span>
                      <button 
                        onClick={() => exportCSV(recentTx, 'recent_transactions')}
                        className="flex items-center gap-1 text-slate-600 hover:text-slate-900 transition text-[11px] font-bold"
                      >
                        <Download className="w-3.5 h-3.5" />
                        <span>تصدير سريع (CSV)</span>
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            } />

            {/* TAB: TRANSACTIONS */}
            <Route path="/transactions" element={
              <div className="space-y-4">
                <div className="enterprise-card rounded-2xl p-6">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4 mb-4">
                    <div>
                      <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                        <ReceiptText className="w-5 h-5 text-blue-600" />
                        <span>سجل القيود والسندات المحاسبية المجمعة (Journal Entries Ledger)</span>
                      </h2>
                      <p className="text-xs text-slate-500 mt-0.5">
                        يتم تجميع كافة الحركات في قيود موحدة حسب رقم القيد (ReqNum). انقر على أي قيد لتوسيع بنوده بالتفصيل بالجنيه والدولار.
                      </p>
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => exportCSV(groupedEntries, 'journal_entries_grouped')}
                        className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-700 px-3.5 py-1.5 rounded-xl text-xs font-bold transition shadow-2xs"
                      >
                        <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                        <span>تصدير القيود المجمعة (Excel)</span>
                      </button>
                      <button
                        onClick={() => window.print()}
                        className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-700 px-3.5 py-1.5 rounded-xl text-xs font-bold transition shadow-2xs"
                      >
                        <Printer className="w-3.5 h-3.5 text-blue-600" />
                        <span>طباعة</span>
                      </button>
                    </div>
                  </div>

                  {/* Filter Toolbar */}
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="relative w-full sm:w-80">
                      <Search className="w-4 h-4 absolute right-3 top-2.5 text-slate-400" />
                      <input
                        type="text"
                        value={entrySearch}
                        onChange={(e) => setEntrySearch(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            setEntryPage(1);
                            loadGroupedEntries();
                          }
                        }}
                        placeholder="ابحث برقم القيد، الحساب، أو البيان..."
                        className="w-full bg-slate-50 border border-slate-300 rounded-xl pr-9 pl-3 py-1.5 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-blue-600 focus:bg-white transition"
                      />
                    </div>

                    <select
                      value={entryDocType}
                      onChange={(e) => {
                        setEntryDocType(e.target.value);
                        setEntryPage(1);
                      }}
                      className="bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 text-xs text-slate-700 font-medium focus:outline-none focus:border-blue-600 focus:bg-white transition"
                    >
                      <option value="">جميع أنواع السندات والقيود</option>
                      {docTypes.map((dt, idx) => (
                        <option key={idx} value={dt.DocType}>{dt.DocType}</option>
                      ))}
                    </select>

                    <button
                      onClick={() => {
                        setEntryPage(1);
                        loadGroupedEntries();
                      }}
                      className="bg-blue-600 hover:bg-blue-700 text-white font-bold px-4 py-1.5 rounded-xl text-xs shadow-xs transition"
                    >
                      تطبيق الفلترة
                    </button>

                    <div className="mr-auto text-xs text-slate-500 font-mono">
                      إجمالي القيود: <span className="font-bold text-slate-900 font-sans">{entryTotal.toLocaleString('en-US')}</span> قيد
                    </div>
                  </div>
                </div>

                {/* Grouped Entries List */}
                <div className="space-y-3">
                  {loadingEntries ? (
                    <div className="enterprise-card rounded-2xl p-12 text-center text-slate-400">
                      جاري تجميع وتجهيز بيانات القيود من SQL Server...
                    </div>
                  ) : groupedEntries.length === 0 ? (
                    <div className="enterprise-card rounded-2xl p-12 text-center text-slate-400">
                      لا توجد قيود مطابقة لمعايير البحث المحددة
                    </div>
                  ) : (
                    groupedEntries.map((entry) => {
                      const isExpanded = !!expandedReqNums[entry.ReqNum];
                      const lines = entryDetailsCache[entry.ReqNum];
                      const isLoadingLines = loadingReqDetails[entry.ReqNum];

                      return (
                        <div 
                          key={entry.ReqNum} 
                          className={`enterprise-card rounded-2xl overflow-hidden transition-all duration-200 ${
                            isExpanded ? 'border-blue-300 ring-2 ring-blue-500/10' : ''
                          }`}
                        >
                          <div 
                            onClick={() => toggleExpandEntry(entry.ReqNum)}
                            className="p-4 cursor-pointer hover:bg-slate-50/70 transition flex flex-col lg:flex-row lg:items-center justify-between gap-4 select-none"
                          >
                            <div className="flex items-center gap-3">
                              <button 
                                className={`w-8 h-8 rounded-lg flex items-center justify-center transition ${
                                  isExpanded ? 'bg-blue-600 text-white shadow-xs' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                                }`}
                              >
                                {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
                              </button>

                              <div>
                                <div className="flex items-center gap-2">
                                  <span className="font-mono font-bold text-blue-700 text-sm">
                                    {entry.ReqNum}
                                  </span>
                                  <span className="bg-slate-100 text-slate-700 px-2 py-0.5 rounded text-[10px] font-semibold border border-slate-200">
                                    {entry.docType || 'قيد'}
                                  </span>
                                  <span className="bg-blue-50 text-blue-800 px-2 py-0.5 rounded text-[10px] font-bold border border-blue-200/80">
                                    {entry.itemsCount} أسطر حركة
                                  </span>
                                </div>
                                <div className="text-xs text-slate-500 mt-1 flex items-center gap-3">
                                  <span>التاريخ: <span className="font-mono font-semibold text-slate-700">{formatDate(entry.theDate)}</span></span>
                                  <span className="text-slate-300">|</span>
                                  <span className="truncate max-w-md font-sans text-slate-600" title={entry.sampleDetail}>
                                    {entry.sampleDetail || 'لا يوجد بيان'}
                                  </span>
                                </div>
                              </div>
                            </div>

                            <div className="flex items-center gap-6 justify-between lg:justify-end">
                              <div className="text-left bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200/80">
                                <div className="text-[10px] font-semibold text-slate-500 uppercase">إجمالي الجنيه (EGP)</div>
                                <div className="font-mono text-xs font-bold text-slate-900 mt-0.5">
                                  <span className="text-rose-700">{formatCurrency(entry.totalDebtor)}</span> / <span className="text-emerald-700">{formatCurrency(entry.totalCreditor)}</span>
                                </div>
                              </div>

                              <div className="text-left bg-blue-50/60 px-3 py-1.5 rounded-xl border border-blue-200/70">
                                <div className="text-[10px] font-semibold text-blue-700 uppercase flex items-center gap-1">
                                  <DollarSign className="w-3 h-3" />
                                  <span>إجمالي الدولار (USD)</span>
                                </div>
                                <div className="font-mono text-xs font-bold text-blue-950 mt-0.5">
                                  ${formatCurrency(entry.totalDebtorDolar)}
                                </div>
                              </div>

                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  openVoucherModal(entry.ReqNum);
                                }}
                                className="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 p-2 rounded-xl transition flex items-center gap-1 font-semibold"
                                title="عرض سند القيد للطباعة"
                              >
                                <Printer className="w-3.5 h-3.5 text-slate-600" />
                                <span className="hidden sm:inline">سند القيد</span>
                              </button>
                            </div>
                          </div>

                          {isExpanded && (
                            <div className="border-t border-slate-200/80 bg-slate-50/50 p-4">
                              {isLoadingLines ? (
                                <div className="py-6 text-center text-xs text-slate-400">
                                  جاري تحميل أسطر القيد...
                                </div>
                              ) : lines && lines.length > 0 ? (
                                <div className="overflow-x-auto">
                                  <table className="w-full text-right text-xs bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
                                    <thead className="bg-slate-100/90 text-slate-600 font-bold border-b border-slate-200">
                                      <tr>
                                        <th className="py-2.5 px-3 font-mono">رقم الحساب</th>
                                        <th className="py-2.5 px-3">اسم الحساب</th>
                                        <th className="py-2.5 px-3">الحساب الرئيسي</th>
                                        <th className="py-2.5 px-3 font-mono text-left bg-rose-50/50 text-rose-800">مدين (جنيه)</th>
                                        <th className="py-2.5 px-3 font-mono text-left bg-emerald-50/50 text-emerald-800">دائن (جنيه)</th>
                                        <th className="py-2.5 px-3 font-mono text-left bg-blue-50/50 text-blue-800">مدين (دولار $)</th>
                                        <th className="py-2.5 px-3 font-mono text-left bg-blue-50/50 text-blue-800">دائن (دولار $)</th>
                                        <th className="py-2.5 px-3">رقم الشيك</th>
                                        <th className="py-2.5 px-3">البيان والشرح</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 font-mono">
                                      {lines.map((line) => (
                                        <tr key={line.Id} className="hover:bg-slate-50/80 transition">
                                          <td className="py-2 px-3 text-slate-500 font-semibold">{line.AccountNumber}</td>
                                          <td className="py-2 px-3 font-sans font-bold text-slate-900">{line.AccountName}</td>
                                          <td className="py-2 px-3 font-sans text-slate-500 text-[11px]">{line.PAccountName || '-'}</td>
                                          <td className="py-2 px-3 text-left font-bold text-rose-700 bg-rose-50/20">
                                            {line.Debtor > 0 ? formatCurrency(line.Debtor) : '-'}
                                          </td>
                                          <td className="py-2 px-3 text-left font-bold text-emerald-700 bg-emerald-50/20">
                                            {line.Creditor > 0 ? formatCurrency(line.Creditor) : '-'}
                                          </td>
                                          <td className="py-2 px-3 text-left font-bold text-blue-700 bg-blue-50/20">
                                            {line.DebtorDolar > 0 ? `$${formatCurrency(line.DebtorDolar)}` : '-'}
                                          </td>
                                          <td className="py-2 px-3 text-left font-bold text-blue-700 bg-blue-50/20">
                                            {line.CreditorDolar > 0 ? `$${formatCurrency(line.CreditorDolar)}` : '-'}
                                          </td>
                                          <td className="py-2 px-3 text-slate-500">{line.Check_Number && line.Check_Number !== '0' ? line.Check_Number : '-'}</td>
                                          <td className="py-2 px-3 font-sans text-slate-600 max-w-xs truncate" title={line.Detail}>
                                            {line.Detail || '-'}
                                          </td>
                                        </tr>
                                      ))}
                                    </tbody>
                                    <tfoot className="bg-slate-100/90 font-mono font-bold text-slate-900 border-t border-slate-200">
                                      <tr>
                                        <td colSpan={3} className="py-2.5 px-3 font-sans text-left">
                                          إجمالي القيد المحاسبي:
                                        </td>
                                        <td className="py-2.5 px-3 text-left text-rose-700 bg-rose-50/40">
                                          {formatCurrency(entry.totalDebtor)} ج.س
                                        </td>
                                        <td className="py-2.5 px-3 text-left text-emerald-700 bg-emerald-50/40">
                                          {formatCurrency(entry.totalCreditor)} ج.س
                                        </td>
                                        <td className="py-2.5 px-3 text-left text-blue-800 bg-blue-50/40">
                                          ${formatCurrency(entry.totalDebtorDolar)}
                                        </td>
                                        <td className="py-2.5 px-3 text-left text-blue-800 bg-blue-50/40">
                                          ${formatCurrency(entry.totalCreditorDolar)}
                                        </td>
                                        <td colSpan={2} className="py-2.5 px-3 font-sans text-emerald-700 text-[11px]">
                                          {Math.abs(entry.totalDebtor - entry.totalCreditor) < 0.1 ? '● القيد متوازن تماماً ✓' : '● يوجد فارق توازن'}
                                        </td>
                                      </tr>
                                    </tfoot>
                                  </table>
                                </div>
                              ) : (
                                <div className="py-4 text-center text-xs text-slate-400">لا توجد أسطر مسجلة</div>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })
                  )}
                </div>

                {/* Pagination */}
                <div className="enterprise-card rounded-2xl p-4 flex items-center justify-between text-xs text-slate-600 font-mono">
                  <div>
                    إجمالي القيود: <span className="font-bold text-slate-900 font-sans">{entryTotal.toLocaleString('en-US')}</span> قيد مسجل
                  </div>

                  <div className="flex items-center gap-3">
                    <button
                      disabled={entryPage <= 1}
                      onClick={() => setEntryPage(p => Math.max(p - 1, 1))}
                      className="p-1.5 rounded-lg bg-slate-50 border border-slate-300 disabled:opacity-40 hover:bg-slate-100 text-slate-700 transition"
                    >
                      <ChevronRight className="w-4 h-4" />
                    </button>
                    <span className="font-sans text-slate-700">
                      صفحة <span className="font-mono text-slate-900 font-bold">{entryPage}</span> من <span className="font-mono text-slate-900 font-bold">{entryTotalPages}</span>
                    </span>
                    <button
                      disabled={entryPage >= entryTotalPages}
                      onClick={() => setEntryPage(p => p + 1)}
                      className="p-1.5 rounded-lg bg-slate-50 border border-slate-300 disabled:opacity-40 hover:bg-slate-100 text-slate-700 transition"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </div>
            } />

            {/* TAB: ACCOUNTS */}
            <Route path="/accounts" element={
              <div className="space-y-4">
                <div className="enterprise-card rounded-2xl p-6">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4 mb-4">
                    <div>
                      <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                        <FolderTree className="w-5 h-5 text-blue-600" />
                        <span>دليل شجرة الحسابات الهرمية التفاعلية (Chart of Accounts Tree)</span>
                      </h2>
                      <p className="text-xs text-slate-500 mt-0.5">
                        عرض هرمي متدرج للحسابات الرئيسية والفرعية مع إمكانية الطي والتوسيع التفاعلي لكافة فروع الحسابات.
                      </p>
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={expandAllTree}
                        className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-700 px-3.5 py-1.5 rounded-xl text-xs font-bold transition shadow-2xs"
                      >
                        <FolderOpen className="w-3.5 h-3.5 text-blue-600" />
                        <span>توسيع الكل</span>
                      </button>
                      <button
                        onClick={collapseAllTree}
                        className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100 border border-slate-300 text-slate-700 px-3.5 py-1.5 rounded-xl text-xs font-bold transition shadow-2xs"
                      >
                        <Folder className="w-3.5 h-3.5 text-slate-500" />
                        <span>طي الكل</span>
                      </button>
                    </div>
                  </div>

                  {/* Primary Accounting Classes Pills */}
                  <div className="flex flex-wrap items-center gap-2 mb-4">
                    <button
                      onClick={() => setSelectedCategory('all')}
                      className={`px-3.5 py-1.5 rounded-xl text-xs font-bold transition ${
                        selectedCategory === 'all'
                          ? 'bg-slate-900 text-white shadow-xs'
                          : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      كافة الأقسام ({rawAccounts.length})
                    </button>
                    <button
                      onClick={() => setSelectedCategory('01')}
                      className={`px-3.5 py-1.5 rounded-xl text-xs font-bold transition flex items-center gap-1.5 ${
                        selectedCategory === '01'
                          ? 'bg-blue-600 text-white shadow-xs'
                          : 'bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100'
                      }`}
                    >
                      <span className="w-2 h-2 rounded-full bg-blue-500"></span>
                      <span>01 - الأصول (Assets)</span>
                    </button>
                    <button
                      onClick={() => setSelectedCategory('02')}
                      className={`px-3.5 py-1.5 rounded-xl text-xs font-bold transition flex items-center gap-1.5 ${
                        selectedCategory === '02'
                          ? 'bg-amber-600 text-white shadow-xs'
                          : 'bg-amber-50 text-amber-700 border border-amber-200 hover:bg-amber-100'
                      }`}
                    >
                      <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                      <span>02 - الخصوم (Liabilities)</span>
                    </button>
                    <button
                      onClick={() => setSelectedCategory('03')}
                      className={`px-3.5 py-1.5 rounded-xl text-xs font-bold transition flex items-center gap-1.5 ${
                        selectedCategory === '03'
                          ? 'bg-rose-600 text-white shadow-xs'
                          : 'bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-100'
                      }`}
                    >
                      <span className="w-2 h-2 rounded-full bg-rose-500"></span>
                      <span>03 - المصروفات (Expenses)</span>
                    </button>
                    <button
                      onClick={() => setSelectedCategory('04')}
                      className={`px-3.5 py-1.5 rounded-xl text-xs font-bold transition flex items-center gap-1.5 ${
                        selectedCategory === '04'
                          ? 'bg-emerald-600 text-white shadow-xs'
                          : 'bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100'
                      }`}
                    >
                      <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                      <span>04 - الإيرادات (Revenues)</span>
                    </button>
                  </div>

                  {/* Filter Toolbar */}
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="relative w-full sm:w-80">
                      <Search className="w-4 h-4 absolute right-3 top-2.5 text-slate-400" />
                      <input
                        type="text"
                        value={accountSearch}
                        onChange={(e) => setAccountSearch(e.target.value)}
                        placeholder="ابحث باسم الحساب أو كوده للتوسيع والتمييز الفوري..."
                        className="w-full bg-slate-50 border border-slate-300 rounded-xl pr-9 pl-3 py-1.5 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-blue-600 focus:bg-white transition"
                      />
                    </div>

                    <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl text-xs font-semibold text-slate-600">
                      <button
                        onClick={() => setSelectedAccountTypeFilter('all')}
                        className={`px-3 py-1 rounded-lg transition ${selectedAccountTypeFilter === 'all' ? 'bg-white shadow-2xs text-slate-900 font-bold' : ''}`}
                      >
                        الكل
                      </button>
                      <button
                        onClick={() => setSelectedAccountTypeFilter('general')}
                        className={`px-3 py-1 rounded-lg transition ${selectedAccountTypeFilter === 'general' ? 'bg-white shadow-2xs text-blue-700 font-bold' : ''}`}
                      >
                        رئيسي (عام)
                      </button>
                      <button
                        onClick={() => setSelectedAccountTypeFilter('detail')}
                        className={`px-3 py-1 rounded-lg transition ${selectedAccountTypeFilter === 'detail' ? 'bg-white shadow-2xs text-slate-900 font-bold' : ''}`}
                      >
                        تحليلي (تفاصيل)
                      </button>
                    </div>

                    {accountSearch && (
                      <button
                        onClick={() => setAccountSearch('')}
                        className="text-xs text-rose-600 hover:underline flex items-center gap-1"
                      >
                        <X className="w-3.5 h-3.5" />
                        <span>إلغاء البحث</span>
                      </button>
                    )}
                  </div>
                </div>

                {/* Main Tree Canvas & Inspector */}
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                  <div className="lg:col-span-2 enterprise-card rounded-2xl p-5 min-h-[500px]">
                    <div className="flex items-center justify-between pb-3 mb-3 border-b border-slate-100">
                      <div className="text-xs font-bold text-slate-700 flex items-center gap-2">
                        <ListTree className="w-4 h-4 text-blue-600" />
                        <span>مستكشف الدليل المحاسبي الهرمي</span>
                      </div>
                      <span className="text-[11px] text-slate-500 font-mono">
                        عرض: <span className="font-bold text-slate-900">{filteredTreeData.length}</span> فروع جذرية
                      </span>
                    </div>

                    {loadingAccounts ? (
                      <div className="py-20 text-center text-slate-400 text-xs">
                        جاري بناء وتنظيم شجرة الحسابات الهرمية...
                      </div>
                    ) : filteredTreeData.length === 0 ? (
                      <div className="py-20 text-center text-slate-400 text-xs">
                        لا توجد حسابات مطابقة لمعايير البحث
                      </div>
                    ) : (
                      <div className="space-y-1 overflow-x-auto">
                        {filteredTreeData.map(rootNode => renderTreeNode(rootNode))}
                      </div>
                    )}
                  </div>

                  <div className="enterprise-card rounded-2xl p-5 flex flex-col justify-between h-fit sticky top-6">
                    <div>
                      <div className="flex items-center justify-between pb-3 mb-4 border-b border-slate-100">
                        <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                          <Tag className="w-4 h-4 text-blue-600" />
                          <span>بطاقة الحساب المحدد</span>
                        </h3>
                        <span className="text-[10px] font-mono text-slate-400 uppercase font-bold">Inspector</span>
                      </div>

                      {selectedAccountNode ? (
                        <div className="space-y-4">
                          <div className="bg-slate-50 p-4 rounded-xl border border-slate-200/80">
                            <span className="text-[11px] font-semibold text-slate-500 block mb-1">كود الحساب:</span>
                            <span className="font-mono text-xl font-bold text-blue-700 block">
                              {selectedAccountNode.code}
                            </span>
                            <span className="text-[11px] font-semibold text-slate-500 block mt-3 mb-1">اسم الحساب:</span>
                            <span className="font-sans font-bold text-base text-slate-900 block">
                              {selectedAccountNode.name}
                            </span>
                          </div>

                          <div className="space-y-2.5 text-xs text-slate-600 bg-white p-3 rounded-xl border border-slate-200">
                            <div className="flex justify-between items-center py-1 border-b border-slate-100">
                              <span className="text-slate-500">التصنيف المحاسبي:</span>
                              <span className="font-bold text-slate-900">{selectedAccountNode.type}</span>
                            </div>
                            <div className="flex justify-between items-center py-1 border-b border-slate-100">
                              <span className="text-slate-500">طبيعة الحساب:</span>
                              <span className="font-semibold text-blue-800">
                                {selectedAccountNode.isGeneral ? 'حساب رئيسي جامع (عــــام)' : 'حساب تحليلي تفصيلي للقيود'}
                              </span>
                            </div>
                            <div className="flex justify-between items-center py-1 border-b border-slate-100">
                              <span className="text-slate-500">الحساب الأب (المباشر):</span>
                              <span className="font-bold text-slate-800 font-mono text-left">{selectedAccountNode.parentName || selectedAccountNode.parentCode}</span>
                            </div>
                            <div className="flex justify-between items-center py-1">
                              <span className="text-slate-500">كود الحساب الأب:</span>
                              <span className="font-mono font-bold text-slate-700">{selectedAccountNode.parentCode}</span>
                            </div>
                          </div>

                          {selectedAccountNode.children.length > 0 && (
                            <div className="bg-blue-50/60 p-3.5 rounded-xl border border-blue-200/80 text-xs text-blue-900">
                              <span className="font-bold block mb-1">الحسابات الفرعية المباشرة:</span>
                              <span className="font-mono text-blue-700 font-bold">{selectedAccountNode.children.length}</span> حساب فرعي تحت هذا المجلد
                            </div>
                          )}

                          <button
                            onClick={() => {
                              navigate(`/transactions?search=${encodeURIComponent(selectedAccountNode.name)}`);
                            }}
                            className="w-full mt-2 bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-3 rounded-xl text-xs transition flex items-center justify-center gap-1.5 shadow-xs"
                          >
                            <ReceiptText className="w-4 h-4" />
                            <span>عرض حركات وقيود هذا الحساب</span>
                          </button>
                        </div>
                      ) : (
                        <div className="py-16 text-center text-slate-400 text-xs">
                          <FolderTree className="w-10 h-10 mx-auto text-slate-300 mb-2 stroke-[1.2]" />
                          <p>انقر على أي حساب من الشجرة لاستعراض بطاقته وتفاصيله الهرمية الكاملة</p>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            } />

            {/* TAB: REPORTS */}
            <Route path="/reports" element={
              <div className="space-y-4">
                {/* REPORT SUB-NAVIGATION TABS */}
                <div className="flex items-center justify-between bg-white p-2 rounded-2xl border border-slate-200/90 shadow-2xs no-print">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setReportTab('generalLedger')}
                      className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                        reportTab === 'generalLedger'
                          ? 'bg-blue-600 text-white shadow-sm shadow-blue-600/30'
                          : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                      }`}
                    >
                      <BookOpen className="w-4 h-4" />
                      <span>تقرير كشف الحسابات العامة (General Ledger)</span>
                      <span className={`text-[10px] px-1.5 py-0.2 rounded font-mono font-bold ${
                        reportTab === 'generalLedger' ? 'bg-blue-700 text-white' : 'bg-blue-50 text-blue-700 border border-blue-200'
                      }`}>
                        عام
                      </span>
                    </button>

                    <button
                      type="button"
                      onClick={() => setReportTab('subAccountLedger')}
                      className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                        reportTab === 'subAccountLedger'
                          ? 'bg-blue-600 text-white shadow-sm shadow-blue-600/30'
                          : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                      }`}
                    >
                      <FileText className="w-4 h-4" />
                      <span>تقرير كشف الحسابات الفرعية (Subsidiary Ledger)</span>
                      <span className={`text-[10px] px-1.5 py-0.2 rounded font-mono font-bold ${
                        reportTab === 'subAccountLedger' ? 'bg-blue-700 text-white' : 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                      }`}>
                        تفصيلي
                      </span>
                    </button>

                    <button
                      type="button"
                      onClick={() => setReportTab('trialBalance')}
                      className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
                        reportTab === 'trialBalance'
                          ? 'bg-blue-600 text-white shadow-sm shadow-blue-600/30'
                          : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                      }`}
                    >
                      <Scale className="w-4 h-4" />
                      <span>تقرير ميزان المراجعة الشامل (Trial Balance)</span>
                      <span className={`text-[10px] px-1.5 py-0.2 rounded font-mono font-bold ${
                        reportTab === 'trialBalance' ? 'bg-blue-700 text-white' : 'bg-slate-100 text-slate-500'
                      }`}>
                        المجاميع والأرصدة
                      </span>
                    </button>
                  </div>

                  <div className="hidden md:flex items-center gap-2 text-xs text-slate-500 font-semibold pl-2">
                    <span>التقارير والقوائم المالية المعتمدة</span>
                  </div>
                </div>

                {reportTab === 'generalLedger' ? (
                  <GeneralLedgerView
                    startDate={glStartDate}
                    setStartDate={setGlStartDate}
                    endDate={glEndDate}
                    setEndDate={setGlEndDate}
                    currency={glCurrency}
                    setCurrency={setGlCurrency}
                    mode={glMode}
                    setMode={setGlMode}
                    generalAccounts={glGeneralAccounts}
                    selectedCodes={glSelectedCodes}
                    toggleAccount={toggleGlAccount}
                    selectAllAccounts={selectAllGlAccounts}
                    clearAllAccounts={clearAllGlAccounts}
                    data={glData}
                    loading={glLoading}
                    onRefresh={loadGeneralLedger}
                    onExportExcel={exportGeneralLedgerExcel}
                    formatCurrency={formatCurrency}
                    renderNetBadge={renderNetBadge}
                    openVoucherModal={openVoucherModal}
                  />
                ) : reportTab === 'subAccountLedger' ? (
                  <SubAccountLedgerView
                    subAccounts={subAccountsList}
                    selectedAccountCode={selectedSubAccountCode}
                    setSelectedAccountCode={setSelectedSubAccountCode}
                    startDate={subStartDate}
                    setStartDate={setSubStartDate}
                    endDate={subEndDate}
                    setEndDate={setSubEndDate}
                    currency={subCurrency}
                    setCurrency={setSubCurrency}
                    loading={subLoading}
                    data={subLedgerData}
                    onRefresh={() => loadSubAccountLedger()}
                    onExportExcel={exportSubAccountExcel}
                    formatCurrency={formatCurrency}
                    renderNetBadge={renderNetBadge}
                    openVoucherModal={openVoucherModal}
                  />
                ) : (
                  <div className="space-y-4">
                    {/* Report Configuration & Filters Bar */}
                    <div className="enterprise-card rounded-2xl p-6 no-print">
                      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-100 pb-4 mb-4">
                        <div>
                          <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                            <Scale className="w-5 h-5 text-blue-600" />
                            <span>تقرير ميزان المراجعة الشامل (Trial Balance Ledger)</span>
                          </h2>
                      <p className="text-xs text-slate-500 mt-0.5">
                        كشف مالي تدقيقي بالأرصدة الافتتاحية وحركات الفترة والأرصدة الختامية مع إمكانية التجميع والتصدير المنسق.
                      </p>
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={exportTrialBalanceExcel}
                        className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-xl text-xs font-bold transition shadow-xs"
                      >
                        <FileSpreadsheet className="w-4 h-4" />
                        <span>تصدير Excel منسق (XLSX)</span>
                      </button>
                      <button
                        onClick={() => window.print()}
                        className="flex items-center gap-1.5 bg-slate-900 hover:bg-slate-800 text-white px-4 py-2 rounded-xl text-xs font-bold transition shadow-xs"
                      >
                        <Printer className="w-4 h-4" />
                        <span>طباعة الميزان المعتمد</span>
                      </button>
                    </div>
                  </div>

                  {/* Period, Currency, Grouping, and Mode Controls Toolbar */}
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-3 pt-1">
                    {/* Date Range: From */}
                    <div>
                      <label className="text-[11px] font-bold text-slate-600 block mb-1.5">
                        من تاريخ الفترة:
                      </label>
                      <div className="relative">
                        <Calendar className="w-4 h-4 absolute right-3 top-2.5 text-slate-400" />
                        <input
                          type="date"
                          value={tbStartDate}
                          onChange={(e) => setTbStartDate(e.target.value)}
                          className="w-full bg-slate-50 border border-slate-300 rounded-xl pr-9 pl-3 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-blue-600 font-mono"
                        />
                      </div>
                    </div>

                    {/* Date Range: To */}
                    <div>
                      <label className="text-[11px] font-bold text-slate-600 block mb-1.5">
                        إلى تاريخ الفترة:
                      </label>
                      <div className="relative">
                        <Calendar className="w-4 h-4 absolute right-3 top-2.5 text-slate-400" />
                        <input
                          type="date"
                          value={tbEndDate}
                          onChange={(e) => setTbEndDate(e.target.value)}
                          className="w-full bg-slate-50 border border-slate-300 rounded-xl pr-9 pl-3 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-blue-600 font-mono"
                        />
                      </div>
                    </div>

                    {/* Grouping Selector (حسب الحساب الرئيسي الكبير) */}
                    <div>
                      <label className="text-[11px] font-bold text-slate-600 block mb-1.5 flex items-center justify-between">
                        <span>تجميع الحسابات (Group By):</span>
                        <Layers className="w-3.5 h-3.5 text-blue-600" />
                      </label>
                      <select
                        value={tbGroupBy}
                        onChange={(e: any) => setTbGroupBy(e.target.value)}
                        className="w-full bg-blue-50/70 border border-blue-300 text-blue-900 rounded-xl px-2.5 py-1.5 text-xs font-bold focus:outline-none focus:border-blue-600 shadow-2xs"
                      >
                        <option value="mainParent">حسب الحساب الرئيسي الكبير (موصى به)</option>
                        <option value="root">حسب الحساب الجذري (أصول، خصوم، مصروفات، إيرادات)</option>
                        <option value="subParent">حسب المجموعة التحليلية الفرعية</option>
                        <option value="none">بدون تجميع (عرض مسطح مباشر)</option>
                      </select>
                    </div>

                    {/* Currency Selector (جنيه / دولار / كلاهما) */}
                    <div>
                      <label className="text-[11px] font-bold text-slate-600 block mb-1.5">
                        عملة التقرير:
                      </label>
                      <div className="flex items-center bg-slate-100 p-1 rounded-xl text-xs font-bold text-slate-600">
                        <button
                          onClick={() => setTbCurrency('EGP')}
                          className={`flex-1 py-1 rounded-lg transition ${tbCurrency === 'EGP' ? 'bg-white shadow-2xs text-blue-700' : ''}`}
                        >
                          جنيه
                        </button>
                        <button
                          onClick={() => setTbCurrency('USD')}
                          className={`flex-1 py-1 rounded-lg transition ${tbCurrency === 'USD' ? 'bg-white shadow-2xs text-blue-700' : ''}`}
                        >
                          دولار ($)
                        </button>
                        <button
                          onClick={() => setTbCurrency('BOTH')}
                          className={`flex-1 py-1 rounded-lg transition ${tbCurrency === 'BOTH' ? 'bg-white shadow-2xs text-blue-700' : ''}`}
                        >
                          كلاهما
                        </button>
                      </div>
                    </div>

                    {/* Mode Selector (تجميعي / تفصيلي) */}
                    <div>
                      <label className="text-[11px] font-bold text-slate-600 block mb-1.5">
                        مستوى البنود:
                      </label>
                      <div className="flex items-center bg-slate-100 p-1 rounded-xl text-xs font-bold text-slate-600">
                        <button
                          onClick={() => setTbMode('detailed')}
                          className={`flex-1 py-1 rounded-lg transition ${tbMode === 'detailed' ? 'bg-white shadow-2xs text-blue-700' : ''}`}
                        >
                          تفصيلي
                        </button>
                        <button
                          onClick={() => setTbMode('summary')}
                          className={`flex-1 py-1 rounded-lg transition ${tbMode === 'summary' ? 'bg-white shadow-2xs text-blue-700' : ''}`}
                        >
                          تجميعي فقط
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Secondary Quick Search & Filter Toolbar */}
                  <div className="flex flex-wrap items-center justify-between gap-3 mt-4 pt-3 border-t border-slate-100">
                    <div className="flex flex-wrap items-center gap-2">
                      <div className="relative w-full sm:w-64">
                        <Search className="w-4 h-4 absolute right-3 top-2 text-slate-400" />
                        <input
                          type="text"
                          value={tbSearch}
                          onChange={(e) => setTbSearch(e.target.value)}
                          placeholder="ابحث بالحساب أو المجموعة في الميزان..."
                          className="w-full bg-slate-50 border border-slate-300 rounded-xl pr-9 pl-3 py-1 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-blue-600"
                        />
                      </div>

                      <select
                        value={tbCategory}
                        onChange={(e) => setTbCategory(e.target.value)}
                        className="bg-slate-50 border border-slate-300 rounded-xl px-3 py-1 text-xs text-slate-700 font-semibold focus:outline-none focus:border-blue-600"
                      >
                        <option value="all">كافة الفئات (أصول، خصوم، مصروفات، إيرادات)</option>
                        <option value="الأصول">الأصول فقط</option>
                        <option value="الخصوم">الخصوم فقط</option>
                        <option value="المصروفات">المصروفات فقط</option>
                        <option value="الإيرادات">الإيرادات فقط</option>
                      </select>

                      <button
                        onClick={loadTrialBalance}
                        className="bg-blue-600 hover:bg-blue-700 text-white font-bold px-3.5 py-1 rounded-xl text-xs shadow-xs transition"
                      >
                        تحديث الفترة
                      </button>
                    </div>

                    <div className="flex items-center gap-2">
                      {tbGroupBy !== 'none' && (
                        <>
                          <button
                            onClick={expandAllGroups}
                            className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold px-2.5 py-1 rounded-lg text-[11px] transition"
                          >
                            توسيع كافة المجموعات
                          </button>
                          <button
                            onClick={collapseAllGroups}
                            className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold px-2.5 py-1 rounded-lg text-[11px] transition"
                          >
                            طي كافة المجموعات
                          </button>
                        </>
                      )}

                      <div className="text-xs text-slate-500 font-mono bg-slate-50 border border-slate-200 px-2.5 py-1 rounded-lg">
                        <span>المجموعات: </span>
                        <span className="font-bold text-slate-900 font-sans">{tbGroupedData.length}</span>
                        <span className="mx-1.5 text-slate-300">|</span>
                        <span>الحسابات: </span>
                        <span className="font-bold text-slate-900 font-sans">{processedTrialBalanceRows.length}</span>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Printable Official Header (Only in Print Mode) */}
                <div className="hidden print:block p-4 border-b border-slate-300 mb-4 text-center">
                  <h1 className="text-xl font-bold text-slate-900">الرويال للتطوير العقاري - الإدارة المالية</h1>
                  <h2 className="text-base font-bold text-slate-700 mt-1">
                    ميزان المراجعة بالأرصدة والمجاميع (مجمع حسب الحساب الرئيسي الكبير)
                  </h2>
                  <p className="text-xs text-slate-500 mt-0.5">
                    عن الفترة من: {tbStartDate} إلى: {tbEndDate} | العملة: {tbCurrency === 'EGP' ? 'جنيه سوداني' : tbCurrency === 'USD' ? 'دولار أمريكي' : 'جنيه ودولار'}
                  </p>
                </div>

                {/* Trial Balance Main Accounting Table */}
                <div className="enterprise-card rounded-2xl p-6 overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-right text-xs">
                      {/* Double-Row Header for Clean Accounting Structure */}
                      <thead className="bg-slate-100/90 text-slate-700 font-bold border-y border-slate-300">
                        <tr>
                          <th rowSpan={2} className="py-2.5 px-3 font-mono border-l border-slate-200">كود الحساب</th>
                          <th rowSpan={2} className="py-2.5 px-3 border-l border-slate-200">اسم الحساب / المجموعة</th>
                          <th rowSpan={2} className="py-2.5 px-3 border-l border-slate-200">التصنيف</th>

                          {/* Opening Balances */}
                          <th colSpan={tbCurrency === 'BOTH' ? 2 : 1} className="py-1.5 px-2 text-center border-l border-slate-200 bg-slate-200/60">
                            الرصيد الافتتاحي السابـق
                          </th>

                          {/* Period Movements */}
                          <th colSpan={tbCurrency === 'BOTH' ? 4 : 2} className="py-1.5 px-2 text-center border-l border-slate-200 bg-blue-100/50 text-blue-900">
                            حركات الفترة المحددة ({tbStartDate} إلى {tbEndDate})
                          </th>

                          {/* Cumulative Ending Balances */}
                          <th colSpan={tbCurrency === 'BOTH' ? 2 : 1} className="py-1.5 px-2 text-center bg-slate-200/60">
                            الرصيد الختامي التراكمـي
                          </th>
                        </tr>
                        <tr className="border-t border-slate-200 text-[11px]">
                          {/* Opening Sub-headers */}
                          {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                            <th className="py-1.5 px-2 text-left font-mono text-slate-800 bg-slate-100 border-l border-slate-200">
                              الرصيد (ج.س)
                            </th>
                          )}
                          {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                            <th className="py-1.5 px-2 text-left font-mono text-blue-900 bg-blue-50/50 border-l border-slate-200">
                              الرصيد ($)
                            </th>
                          )}

                          {/* Period Movements Sub-headers */}
                          {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                            <>
                              <th className="py-1.5 px-2 text-left font-mono text-rose-800 bg-rose-50/50">مدين ج.س</th>
                              <th className="py-1.5 px-2 text-left font-mono text-emerald-800 bg-emerald-50/50 border-l border-slate-200">دائن ج.س</th>
                            </>
                          )}
                          {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                            <>
                              <th className="py-1.5 px-2 text-left font-mono text-blue-800 bg-blue-50/50">مدين $</th>
                              <th className="py-1.5 px-2 text-left font-mono text-blue-800 bg-blue-50/50 border-l border-slate-200">دائن $</th>
                            </>
                          )}

                          {/* Ending Sub-headers */}
                          {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                            <th className="py-1.5 px-2 text-left font-mono text-slate-800 bg-slate-100 border-l border-slate-200">
                              الرصيد (ج.س)
                            </th>
                          )}
                          {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                            <th className="py-1.5 px-2 text-left font-mono text-blue-900 bg-blue-50/50">
                              الرصيد ($)
                            </th>
                          )}
                        </tr>
                      </thead>

                      {/* Body with Group By Hierarchy */}
                      <tbody className="divide-y divide-slate-100 font-mono">
                        {tbLoading ? (
                          <tr>
                            <td colSpan={tbCurrency === 'BOTH' ? 11 : 7} className="py-16 text-center text-slate-400 font-sans">
                              جاري معالجة وتجميع أرصدة ميزان المراجعة من قاعدة البيانات...
                            </td>
                          </tr>
                        ) : tbGroupedData.length === 0 ? (
                          <tr>
                            <td colSpan={tbCurrency === 'BOTH' ? 11 : 7} className="py-16 text-center text-slate-400 font-sans">
                              لا توجد حركات مسجلة مطابقة لمعايير البحث والتاريخ
                            </td>
                          </tr>
                        ) : (
                          tbGroupedData.map((group) => {
                            const isCollapsed = collapsedGroups[group.groupKey] !== false;
                            return (
                              <React.Fragment key={group.groupKey}>
                                {/* GROUP HEADER ROW (Group by Big Main Parent) */}
                                <tr 
                                  onClick={() => toggleGroup(group.groupKey)}
                                  className="bg-slate-100/95 hover:bg-slate-200/90 cursor-pointer transition border-t-2 border-slate-300 font-bold select-none group"
                                  title="انقر لتوسيع أو طي حسابات هذه المجموعة الرئيسية"
                                >
                                  <td colSpan={3} className="py-2.5 px-3">
                                    <div className="flex items-center gap-2">
                                      <span className="p-1 rounded-md bg-white border border-slate-300 text-slate-600 shadow-2xs group-hover:border-blue-500 group-hover:text-blue-600 transition">
                                        {isCollapsed ? <ChevronLeft className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                                      </span>
                                      <span className="font-mono bg-blue-600 text-white px-2 py-0.5 rounded text-[11px] font-bold shadow-2xs">
                                        {group.groupCode}
                                      </span>
                                      <span className="font-sans text-xs text-slate-900 font-extrabold">
                                        {group.groupName}
                                      </span>
                                      <span className="bg-slate-200 text-slate-700 text-[10px] px-2 py-0.5 rounded-full font-sans font-semibold">
                                        {group.category}
                                      </span>
                                      <span className="text-slate-500 font-mono text-[11px] font-normal mr-1">
                                        ({group.rows.length} حساب فرعي)
                                      </span>
                                    </div>
                                  </td>

                                  {/* Opening Net Balances Subtotals */}
                                  {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                                    <td className="py-2.5 px-2 text-left bg-slate-200/50 border-l border-slate-300">
                                      {renderNetBadge(group.totals.openDebtor - group.totals.openCreditor, false, 'ج.س')}
                                    </td>
                                  )}
                                  {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                                    <td className="py-2.5 px-2 text-left bg-blue-100/40 border-l border-slate-300">
                                      {renderNetBadge(group.totals.openDebtorDolar - group.totals.openCreditorDolar, false, '$')}
                                    </td>
                                  )}

                                  {/* Period Movements Subtotals */}
                                  {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                                    <>
                                      <td className="py-2.5 px-2 text-left text-rose-800 bg-rose-100/40 font-bold">
                                        {group.totals.periodDebtor > 0 ? formatCurrency(group.totals.periodDebtor) : '-'}
                                      </td>
                                      <td className="py-2.5 px-2 text-left text-emerald-800 bg-emerald-100/40 font-bold border-l border-slate-300">
                                        {group.totals.periodCreditor > 0 ? formatCurrency(group.totals.periodCreditor) : '-'}
                                      </td>
                                    </>
                                  )}
                                  {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                                    <>
                                      <td className="py-2.5 px-2 text-left text-blue-800 bg-blue-100/40 font-bold">
                                        {group.totals.periodDebtorDolar > 0 ? `$${formatCurrency(group.totals.periodDebtorDolar)}` : '-'}
                                      </td>
                                      <td className="py-2.5 px-2 text-left text-blue-800 bg-blue-100/40 font-bold border-l border-slate-300">
                                        {group.totals.periodCreditorDolar > 0 ? `$${formatCurrency(group.totals.periodCreditorDolar)}` : '-'}
                                      </td>
                                    </>
                                  )}

                                  {/* Ending Cumulative Net Subtotals */}
                                  {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                                    <td className="py-2.5 px-2 text-left bg-slate-200/70 border-l border-slate-300">
                                      {renderNetBadge(group.totals.totalDebtor - group.totals.totalCreditor, true, 'ج.س')}
                                    </td>
                                  )}
                                  {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                                    <td className="py-2.5 px-2 text-left bg-blue-100/50">
                                      {renderNetBadge(group.totals.totalDebtorDolar - group.totals.totalCreditorDolar, true, '$')}
                                    </td>
                                  )}
                                </tr>

                                {/* CHILD DETAIL ACCOUNTS (Collapsed by default on screen, printed fully on print) */}
                                {tbMode === 'detailed' && group.rows.map((r, rIdx) => (
                                  <tr 
                                    key={`${group.groupKey}-row-${rIdx}`} 
                                    className={`${isCollapsed ? 'hidden print:table-row' : ''} hover:bg-blue-50/40 transition select-none`}
                                  >
                                    <td className="py-2 px-3 font-bold text-blue-700 whitespace-nowrap pr-8">
                                      <span className="text-slate-300 font-normal ml-1 font-mono">└─</span>
                                      <span>{r.AccountNumber}</span>
                                    </td>
                                    <td className="py-2 px-3 font-sans font-bold text-slate-800 whitespace-nowrap">
                                      <div className="flex items-center gap-1.5">
                                        <span>{r.AccountName}</span>
                                        {r.SubParentName && r.SubParentName !== group.groupName && (
                                          <span className="text-[10px] bg-slate-100 text-slate-500 px-1.5 py-0.2 rounded font-normal font-sans border border-slate-200">
                                            {r.SubParentName}
                                          </span>
                                        )}
                                      </div>
                                    </td>
                                    <td className="py-2 px-3 font-sans text-slate-500 text-[11px] whitespace-nowrap">
                                      {r.category}
                                    </td>

                                    {/* Opening Net Balances */}
                                    {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                                      <td className="py-2 px-2 text-left bg-slate-50/30 border-l border-slate-200">
                                        {renderNetBadge(r.netOpen, false, 'ج.س')}
                                      </td>
                                    )}
                                    {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                                      <td className="py-2 px-2 text-left bg-blue-50/10 border-l border-slate-200">
                                        {renderNetBadge(r.netOpenDolar, false, '$')}
                                      </td>
                                    )}

                                    {/* Period Movements */}
                                    {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                                      <>
                                        <td className="py-2 px-2 text-left font-bold text-rose-700 bg-rose-50/20">
                                          {r.periodDebtor > 0 ? formatCurrency(r.periodDebtor) : '-'}
                                        </td>
                                        <td className="py-2 px-2 text-left font-bold text-emerald-700 bg-emerald-50/20 border-l border-slate-200">
                                          {r.periodCreditor > 0 ? formatCurrency(r.periodCreditor) : '-'}
                                        </td>
                                      </>
                                    )}
                                    {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                                      <>
                                        <td className="py-2 px-2 text-left font-bold text-blue-700 bg-blue-50/20">
                                          {r.periodDebtorDolar > 0 ? `$${formatCurrency(r.periodDebtorDolar)}` : '-'}
                                        </td>
                                        <td className="py-2 px-2 text-left font-bold text-blue-700 bg-blue-50/20 border-l border-slate-200">
                                          {r.periodCreditorDolar > 0 ? `$${formatCurrency(r.periodCreditorDolar)}` : '-'}
                                        </td>
                                      </>
                                    )}

                                    {/* Ending Net Cumulative */}
                                    {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                                      <td className="py-2 px-2 text-left bg-slate-100/50 border-l border-slate-200">
                                        {renderNetBadge(r.netEnding, true, 'ج.س')}
                                      </td>
                                    )}
                                    {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                                      <td className="py-2 px-2 text-left bg-blue-50/40">
                                        {renderNetBadge(r.netEndingDolar, true, '$')}
                                      </td>
                                    )}
                                  </tr>
                                ))}
                              </React.Fragment>
                            );
                          })
                        )}
                      </tbody>

                      {/* GRAND TOTALS & EQUILIBRIUM FOOTER ROW */}
                      <tfoot className="bg-slate-200/90 font-mono font-bold text-slate-900 border-t-2 border-slate-400 text-xs">
                        <tr>
                          <td colSpan={3} className="py-3 px-3 font-sans text-sm text-slate-900">
                            المجموع العام لميزان المراجعة:
                          </td>

                          {/* Opening Totals: Debits vs Credits */}
                          {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                            <td className="py-2.5 px-2 text-left text-slate-800 border-l border-slate-300">
                              <div className="flex flex-col gap-0.5 text-[11px] font-sans">
                                <div className="flex items-center justify-between gap-1">
                                  <span className="text-[10px] bg-rose-100 text-rose-800 px-1 py-0.2 rounded font-bold">مدين</span>
                                  <span className="font-mono text-rose-700 font-bold">{formatCurrency(tbGrandTotals.netOpenDebit)}</span>
                                </div>
                                <div className="flex items-center justify-between gap-1">
                                  <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1 py-0.2 rounded font-bold">دائن</span>
                                  <span className="font-mono text-emerald-700 font-bold">{formatCurrency(tbGrandTotals.netOpenCredit)}</span>
                                </div>
                              </div>
                            </td>
                          )}
                          {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                            <td className="py-2.5 px-2 text-left text-blue-900 border-l border-slate-300">
                              <div className="flex flex-col gap-0.5 text-[11px] font-sans">
                                <div className="flex items-center justify-between gap-1">
                                  <span className="text-[10px] bg-rose-100 text-rose-800 px-1 py-0.2 rounded font-bold">مدين</span>
                                  <span className="font-mono text-blue-900 font-bold">${formatCurrency(tbGrandTotals.netOpenDebitDolar)}</span>
                                </div>
                                <div className="flex items-center justify-between gap-1">
                                  <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1 py-0.2 rounded font-bold">دائن</span>
                                  <span className="font-mono text-blue-900 font-bold">${formatCurrency(tbGrandTotals.netOpenCreditDolar)}</span>
                                </div>
                              </div>
                            </td>
                          )}

                          {/* Period Movements Totals */}
                          {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                            <>
                              <td className="py-3 px-2 text-left text-rose-800 bg-rose-100/40">
                                {formatCurrency(tbGrandTotals.periodDebtor)}
                              </td>
                              <td className="py-3 px-2 text-left text-emerald-800 bg-emerald-100/40 border-l border-slate-300">
                                {formatCurrency(tbGrandTotals.periodCreditor)}
                              </td>
                            </>
                          )}
                          {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                            <>
                              <td className="py-3 px-2 text-left text-blue-900 bg-blue-100/40">
                                ${formatCurrency(tbGrandTotals.periodDebtorDolar)}
                              </td>
                              <td className="py-3 px-2 text-left text-blue-900 bg-blue-100/40 border-l border-slate-300">
                                ${formatCurrency(tbGrandTotals.periodCreditorDolar)}
                              </td>
                            </>
                          )}

                          {/* Ending Totals: Debits vs Credits */}
                          {(tbCurrency === 'EGP' || tbCurrency === 'BOTH') && (
                            <td className="py-2.5 px-2 text-left text-slate-900 bg-slate-300/40 border-l border-slate-300">
                              <div className="flex flex-col gap-0.5 text-[11px] font-sans">
                                <div className="flex items-center justify-between gap-1">
                                  <span className="text-[10px] bg-rose-100 text-rose-800 px-1 py-0.2 rounded font-bold">مدين</span>
                                  <span className="font-mono text-rose-900 font-extrabold">{formatCurrency(tbGrandTotals.netEndingDebit)}</span>
                                </div>
                                <div className="flex items-center justify-between gap-1">
                                  <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1 py-0.2 rounded font-bold">دائن</span>
                                  <span className="font-mono text-emerald-900 font-extrabold">{formatCurrency(tbGrandTotals.netEndingCredit)}</span>
                                </div>
                              </div>
                            </td>
                          )}
                          {(tbCurrency === 'USD' || tbCurrency === 'BOTH') && (
                            <td className="py-2.5 px-2 text-left text-blue-950 bg-blue-100/50">
                              <div className="flex flex-col gap-0.5 text-[11px] font-sans">
                                <div className="flex items-center justify-between gap-1">
                                  <span className="text-[10px] bg-rose-100 text-rose-800 px-1 py-0.2 rounded font-bold">مدين</span>
                                  <span className="font-mono text-blue-950 font-extrabold">${formatCurrency(tbGrandTotals.netEndingDebitDolar)}</span>
                                </div>
                                <div className="flex items-center justify-between gap-1">
                                  <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1 py-0.2 rounded font-bold">دائن</span>
                                  <span className="font-mono text-blue-950 font-extrabold">${formatCurrency(tbGrandTotals.netEndingCreditDolar)}</span>
                                </div>
                              </div>
                            </td>
                          )}
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  {/* Print Signatures Block (Only appears on print) */}
                  <div className="hidden print:grid grid-cols-3 gap-8 mt-12 pt-6 border-t border-slate-400 text-center text-xs">
                    <div>
                      <span className="font-bold block">إعداد المحاسب المسؤول</span>
                      <span className="mt-8 block border-b border-dashed border-slate-400 w-40 mx-auto"></span>
                    </div>
                    <div>
                      <span className="font-bold block">تدقيق رئيس الحسابات</span>
                      <span className="mt-8 block border-b border-dashed border-slate-400 w-40 mx-auto"></span>
                    </div>
                    <div>
                      <span className="font-bold block">اعتماد المدير المالي</span>
                      <span className="mt-8 block border-b border-dashed border-slate-400 w-40 mx-auto"></span>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        } />

            <Route path="*" element={<Navigate to="/dashboard" replace />} />
          </Routes>
        </main>
      </div>

      {/* FULL PRINTABLE ACCOUNTING VOUCHER MODAL */}
      {modalEntry && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-300 rounded-2xl w-full max-w-4xl overflow-hidden shadow-2xl flex flex-col max-h-[92vh]">
            <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-blue-600 text-white flex items-center justify-center font-bold shadow-md shadow-blue-500/20">
                  <ReceiptText className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="font-bold text-slate-900 text-base">سند قيد محاسبي رسمي</h4>
                  <div className="flex items-center gap-3 text-xs text-slate-500 mt-0.5">
                    <span>رقم القيد: <span className="font-mono font-bold text-blue-700">{modalEntry.reqNum}</span></span>
                    <span>•</span>
                    <span>التاريخ: <span className="font-mono text-slate-700">{modalEntry.lines[0]?.TheDate ? formatDate(modalEntry.lines[0].TheDate) : '-'}</span></span>
                    <span>•</span>
                    <span className="bg-slate-200/80 px-2 py-0.5 rounded text-[11px] text-slate-800 font-semibold">{modalEntry.lines[0]?.DocType || 'قيد'}</span>
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => window.print()}
                  className="flex items-center gap-1.5 bg-white hover:bg-slate-100 border border-slate-300 text-slate-700 px-3 py-1.5 rounded-xl text-xs font-bold transition shadow-2xs"
                >
                  <Printer className="w-3.5 h-3.5 text-blue-600" />
                  <span>طباعة السند</span>
                </button>
                <button
                  onClick={() => setModalEntry(null)}
                  className="text-slate-400 hover:text-slate-700 p-1.5 rounded-lg hover:bg-slate-200 transition"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            <div className="p-6 overflow-y-auto space-y-4">
              <div className="overflow-x-auto">
                <table className="w-full text-right text-xs border border-slate-200 rounded-xl overflow-hidden">
                  <thead className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                    <tr>
                      <th className="py-2.5 px-3 font-mono">رقم الحساب</th>
                      <th className="py-2.5 px-3">اسم الحساب</th>
                      <th className="py-2.5 px-3 font-mono text-left bg-rose-50 text-rose-800">مدين (جنيه)</th>
                      <th className="py-2.5 px-3 font-mono text-left bg-emerald-50 text-emerald-800">دائن (جنيه)</th>
                      <th className="py-2.5 px-3 font-mono text-left bg-blue-50 text-blue-800">مدين ($)</th>
                      <th className="py-2.5 px-3 font-mono text-left bg-blue-50 text-blue-800">دائن ($)</th>
                      <th className="py-2.5 px-3">البيان والشرح</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono">
                    {modalEntry.lines.map((l) => (
                      <tr key={l.Id} className="hover:bg-slate-50 transition">
                        <td className="py-2.5 px-3 text-slate-600 font-semibold">{l.AccountNumber}</td>
                        <td className="py-2.5 px-3 font-sans font-bold text-slate-900">{l.AccountName}</td>
                        <td className="py-2.5 px-3 text-left font-bold text-rose-700 bg-rose-50/20">
                          {l.Debtor > 0 ? formatCurrency(l.Debtor) : '-'}
                        </td>
                        <td className="py-2.5 px-3 text-left font-bold text-emerald-700 bg-emerald-50/20">
                          {l.Creditor > 0 ? formatCurrency(l.Creditor) : '-'}
                        </td>
                        <td className="py-2.5 px-3 text-left font-bold text-blue-700 bg-blue-50/20">
                          {l.DebtorDolar > 0 ? `$${formatCurrency(l.DebtorDolar)}` : '-'}
                        </td>
                        <td className="py-2.5 px-3 text-left font-bold text-blue-700 bg-blue-50/20">
                          {l.CreditorDolar > 0 ? `$${formatCurrency(l.CreditorDolar)}` : '-'}
                        </td>
                        <td className="py-2.5 px-3 font-sans text-slate-700 max-w-xs">{l.Detail || '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot className="bg-slate-100 font-mono font-bold text-slate-900 border-t border-slate-200">
                    <tr>
                      <td colSpan={2} className="py-3 px-3 font-sans text-left text-sm">
                        الإجمالي العام للقيد:
                      </td>
                      <td className="py-3 px-3 text-left text-rose-700 text-sm bg-rose-50/40">
                        {formatCurrency(modalEntry.lines.reduce((s, x) => s + (x.Debtor || 0), 0))} ج.س
                      </td>
                      <td className="py-3 px-3 text-left text-emerald-700 text-sm bg-emerald-50/40">
                        {formatCurrency(modalEntry.lines.reduce((s, x) => s + (x.Creditor || 0), 0))} ج.س
                      </td>
                      <td className="py-3 px-3 text-left text-blue-800 text-sm bg-blue-50/40">
                        ${formatCurrency(modalEntry.lines.reduce((s, x) => s + (x.DebtorDolar || 0), 0))}
                      </td>
                      <td className="py-3 px-3 text-left text-blue-800 text-sm bg-blue-50/40">
                        ${formatCurrency(modalEntry.lines.reduce((s, x) => s + (x.CreditorDolar || 0), 0))}
                      </td>
                      <td className="py-3 px-3 text-emerald-700 font-sans text-xs">
                        ● القيد متوازن
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
              <button
                onClick={() => setModalEntry(null)}
                className="bg-slate-800 hover:bg-slate-900 text-white font-bold px-5 py-2 rounded-xl text-xs transition shadow-xs"
              >
                إغلاق النافذة
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
