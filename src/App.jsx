import React, { useState, useEffect, useMemo } from 'react';
import {
  IndianRupee,
  Calendar,
  PlusCircle,
  RefreshCw,
  Trash2,
  Edit3,
  Search,
  PieChart as PieChartIcon,
  BarChart2,
  ShieldCheck,
  ArrowRight,
  CheckCircle2,
  XCircle,
  Link2,
  Copy,
  Code,
  HelpCircle,
  X,
  Filter,
  Check,
  ChevronRight,
  DollarSign,
  CreditCard,
  Tag,
  FileText,
  Clock,
  Sparkles,
  Download,
  Moon,
  Sun,
  TrendingUp,
  TrendingDown,
  Target,
  Zap,
  Wand2,
  Smartphone,
  QrCode,
  Globe,
  Share2,
  ExternalLink,
  UploadCloud
} from 'lucide-react';

const CATEGORIES = [
  'Food',
  'Travel',
  'Shopping',
  'Bills',
  'Education',
  'Health',
  'Entertainment',
  'Rent',
  'Groceries',
  'Income / Salary',
  'Other'
];

const PAYMENT_METHODS = [
  'UPI',
  'Cash',
  'Credit Card',
  'Debit Card',
  'Bank Transfer'
];

const QUICK_PRESETS = [
  { label: '☕ Chai / Coffee', category: 'Food', amount: 40, description: 'Tea / Coffee break', paymentMethod: 'UPI' },
  { label: '🍛 Canteen Lunch', category: 'Food', amount: 150, description: 'Lunch at canteen', paymentMethod: 'UPI' },
  { label: '🛺 Auto / Cab', category: 'Travel', amount: 20, description: 'Auto fare to college', paymentMethod: 'UPI' },
  { label: '🛒 Groceries', category: 'Groceries', amount: 650, description: 'Weekly essentials', paymentMethod: 'UPI' },
  { label: '📶 Wi-Fi Bill', category: 'Bills', amount: 999, description: 'Broadband recharge', paymentMethod: 'Credit Card' }
];

const DEFAULT_WEB_APP_URL = "";

const DEFAULT_EXPENSES = [
  {
    rowIndex: 2,
    date: '2026-10-01',
    category: 'Food',
    description: 'Lunch at college canteen',
    amount: 150,
    type: 'Expense',
    paymentMethod: 'UPI',
    notes: 'South Indian Combo',
    createdAt: '2026-10-01T13:20:45.000Z'
  },
  {
    rowIndex: 3,
    date: '2026-10-01',
    category: 'Travel',
    description: 'Auto fare to college',
    amount: 20,
    type: 'Expense',
    paymentMethod: 'UPI',
    notes: 'Shared auto',
    createdAt: '2026-10-01T09:15:00.000Z'
  },
  {
    rowIndex: 4,
    date: '2026-09-28',
    category: 'Shopping',
    description: 'New denim jacket',
    amount: 2499,
    type: 'Expense',
    paymentMethod: 'Credit Card',
    notes: 'Festival discount offer',
    createdAt: '2026-09-28T18:40:10.000Z'
  },
  {
    rowIndex: 5,
    date: '2026-09-25',
    category: 'Bills',
    description: 'High-speed Wi-Fi recharge',
    amount: 999,
    type: 'Expense',
    paymentMethod: 'UPI',
    notes: 'Monthly plan',
    createdAt: '2026-09-25T11:00:00.000Z'
  },
  {
    rowIndex: 6,
    date: '2026-09-20',
    category: 'Groceries',
    description: 'Weekly supermarket stash',
    amount: 1850,
    type: 'Expense',
    paymentMethod: 'Debit Card',
    notes: 'Vegetables & dairy items',
    createdAt: '2026-09-20T17:30:00.000Z'
  },
  {
    rowIndex: 7,
    date: '2026-09-01',
    category: 'Income / Salary',
    description: 'Monthly Stipend / Allowance',
    amount: 15000,
    type: 'Income',
    paymentMethod: 'Bank Transfer',
    notes: 'Monthly allowance credited',
    createdAt: '2026-09-01T08:00:00.000Z'
  }
];

const APPS_SCRIPT_CODE = `function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    rows.push({
      rowIndex: i + 1,
      date: row[0] ? String(row[0]).substring(0,10) : "",
      category: row[1] || "",
      description: row[2] || "",
      amount: Number(row[3]) || 0,
      paymentMethod: row[4] || "",
      notes: row[5] || "",
      createdAt: row[6] || ""
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Category", "Description", "Amount (INR)", "Payment Method", "Notes", "Created At"]);
  }
  
  var body = JSON.parse(e.postData.contents);
  var action = body.action;
  
  if (action === "ADD") {
    sheet.appendRow([
      body.date,
      body.category,
      body.description,
      body.amount,
      body.paymentMethod,
      body.notes || "",
      new Date().toISOString()
    ]);
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "DELETE") {
    sheet.deleteRow(body.rowIndex);
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "EDIT") {
    var range = sheet.getRange(body.rowIndex, 1, 1, 6);
    range.setValues([[
      body.date,
      body.category,
      body.description,
      body.amount,
      body.paymentMethod,
      body.notes || ""
    ]]);
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid action" }))
    .setMimeType(ContentService.MimeType.JSON);
}`;

const safeCopyToClipboard = async (text) => {
  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch (err) {
    console.warn('Clipboard API call failed or blocked by permissions policy, using fallback:', err);
  }

  // Fallback for sandboxed or policy-restricted environments
  try {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.left = '-999999px';
    textArea.style.top = '-999999px';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    const successful = document.execCommand('copy');
    document.body.removeChild(textArea);
    return successful;
  } catch (fallbackErr) {
    console.error('Fallback copy command failed:', fallbackErr);
    return false;
  }
};

export default function App() {
  const [sheetUrl, setSheetUrl] = useState(() => {
    if (typeof window !== 'undefined') {
      const urlParams = new URLSearchParams(window.location.search);
      const queryUrl = urlParams.get('url') || urlParams.get('scriptUrl');
      if (queryUrl) return queryUrl;
      const saved = localStorage.getItem('EXPENSE_TRACKER_SHEET_URL');
      if (saved) return saved;
    }
    return DEFAULT_WEB_APP_URL;
  });

  const [darkMode, setDarkMode] = useState(() => {
    if (typeof window !== 'undefined') {
      return localStorage.getItem('EXPENSE_TRACKER_THEME') === 'dark';
    }
    return false;
  });

  const [monthlyBudget, setMonthlyBudget] = useState(() => {
  const savedBudget = localStorage.getItem('EXPENSE_TRACKER_MONTHLY_BUDGET');
  return savedBudget ? Number(savedBudget) : 12000;
});
  const [isEditingBudget, setIsEditingBudget] = useState(false);

  const [isConnected, setIsConnected] = useState(false);
  const [expenses, setExpenses] = useState(DEFAULT_EXPENSES);
  const [loading, setLoading] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [notification, setNotification] = useState(null);
  const [showCodeModal, setShowCodeModal] = useState(false);
  const [copiedCode, setCopiedCode] = useState(false);
  const [copiedShareLink, setCopiedShareLink] = useState(false);

  // Form states
  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
  const [category, setCategory] = useState('Food');
  const [description, setDescription] = useState('');
  const [amount, setAmount] = useState('');
  const [txType, setTxType] = useState('Expense');
  const [paymentMethod, setPaymentMethod] = useState('UPI');
  const [notes, setNotes] = useState('');

  // AI Smart Fill State
  const [aiPrompt, setAiPrompt] = useState('');
  const [isAiProcessing, setIsAiProcessing] = useState(false);

  // Editing state
  const [editingRowIndex, setEditingRowIndex] = useState(null);

  // Filters
  const [searchQuery, setSearchQuery] = useState('');
  const [filterCategory, setFilterCategory] = useState('All');
  const [filterPayment, setFilterPayment] = useState('All');
  const [filterType, setFilterType] = useState('All');

  // Modal Delete state
  const [deleteTargetIndex, setDeleteTargetIndex] = useState(null);

  // Publish & Mobile Access Modal state
  const [showPublishModal, setShowPublishModal] = useState(false);

  // System Notification Helper
  const notify = (msg, type = 'success') => {
    setNotification({ msg, type });
    setTimeout(() => {
      setNotification(null);
    }, 4500);
  };

  // Google Sheets API Data Fetcher
  const fetchExpenses = async (urlToUse = sheetUrl, isInitial = false) => {
    if (!urlToUse || !urlToUse.startsWith('https://script.google.com/')) {
      if (!isInitial) notify('Please enter a valid Google Apps Script Web App URL.', 'error');
      return;
    }
    setLoading(true);
    try {
      const res = await fetch(urlToUse);
      const data = await res.json();
      if (data.status === 'success') {
        setExpenses(data.data || []);
        setIsConnected(true);
        localStorage.setItem('EXPENSE_TRACKER_SHEET_URL', urlToUse);
        if (!isInitial) notify('Successfully connected and loaded Google Sheet data!', 'success');
      } else {
        setIsConnected(false);
        notify('Google Sheet returned an unexpected response format.', 'error');
      }
    } catch (err) {
      console.error(err);
      setIsConnected(false);
      notify('Failed to connect to Google Sheets API. Please verify deployment settings.', 'error');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (sheetUrl && sheetUrl.startsWith('https://script.google.com/')) {
      fetchExpenses(sheetUrl, true);
    }
  }, []);

  const handleConnect = (e) => {
    e.preventDefault();
    const cleanUrl = sheetUrl.trim();
    if (!cleanUrl.startsWith('https://script.google.com/')) {
      notify('Please enter a valid Google Apps Script Web App URL.', 'error');
      return;
    }
    fetchExpenses(cleanUrl, false);
  };

  const copyShareLink = async () => {
    if (!sheetUrl) return;
    const shareableUrl = `${window.location.origin}${window.location.pathname}?url=${encodeURIComponent(sheetUrl)}`;
    const success = await safeCopyToClipboard(shareableUrl);
    if (success) {
      setCopiedShareLink(true);
      notify('Shareable auto-sync link copied to clipboard!', 'success');
      setTimeout(() => setCopiedShareLink(false), 3000);
    } else {
      notify('Could not copy automatically due to browser restrictions.', 'error');
    }
  };

  const handleAiSmartFill = () => {
    if (!aiPrompt.trim()) return;
    setIsAiProcessing(true);

    setTimeout(() => {
      const text = aiPrompt.toLowerCase();
      let detectedAmount = '';
      let detectedCategory = 'Food';
      let detectedPayment = 'UPI';

      const numMatch = text.match(/(\d+(\.\d+)?)/);
      if (numMatch) detectedAmount = numMatch[0];

      if (text.includes('cab') || text.includes('uber') || text.includes('auto') || text.includes('travel') || text.includes('flight') || text.includes('train')) {
        detectedCategory = 'Travel';
      } else if (text.includes('grocery') || text.includes('vegetable') || text.includes('milk') || text.includes('mart')) {
        detectedCategory = 'Groceries';
      } else if (text.includes('recharge') || text.includes('wifi') || text.includes('bill') || text.includes('electricity')) {
        detectedCategory = 'Bills';
      } else if (text.includes('shirt') || text.includes('jacket') || text.includes('shoe') || text.includes('amazon') || text.includes('shopping')) {
        detectedCategory = 'Shopping';
      } else if (text.includes('movie') || text.includes('game') || text.includes('netflix') || text.includes('concert')) {
        detectedCategory = 'Entertainment';
      } else if (text.includes('salary') || text.includes('stipend') || text.includes('income')) {
        detectedCategory = 'Income / Salary';
      }

      if (text.includes('cash')) detectedPayment = 'Cash';
      else if (text.includes('card') || text.includes('credit')) detectedPayment = 'Credit Card';
      else if (text.includes('debit')) detectedPayment = 'Debit Card';
      else if (text.includes('bank') || text.includes('transfer')) detectedPayment = 'Bank Transfer';

      setDescription(aiPrompt);
      if (detectedAmount) setAmount(detectedAmount);
      setCategory(detectedCategory);
      setPaymentMethod(detectedPayment);
      if (detectedCategory === 'Income / Salary') setTxType('Income');
      else setTxType('Expense');

      setIsAiProcessing(false);
      setAiPrompt('');
      notify('Parsed details automatically into the form!', 'success');
    }, 400);
  };

  const handleApplyPreset = (preset) => {
    setCategory(preset.category);
    setAmount(preset.amount.toString());
    setDescription(preset.description);
    setPaymentMethod(preset.paymentMethod);
    setTxType('Expense');
    notify(`Applied preset: ${preset.label}`, 'success');
  };

  const handleSubmitExpense = async (e) => {
    e.preventDefault();

    const numericAmount = parseFloat(amount);
    if (isNaN(numericAmount) || numericAmount <= 0) {
      notify('Amount must be a positive number greater than 0.', 'error');
      return;
    }
    if (!description.trim()) {
      notify('Description is required.', 'error');
      return;
    }

    setSubmitting(true);
    const finalCategory = txType === 'Income' ? 'Income / Salary' : category;

    const payload = {
      action: editingRowIndex ? 'EDIT' : 'ADD',
      rowIndex: editingRowIndex,
      date,
      category: finalCategory,
      description,
      amount: numericAmount,
      paymentMethod,
      notes
    };

    if (isConnected && sheetUrl) {
      try {
        const res = await fetch(sheetUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(payload)
        });
        const data = await res.json();

        if (data.status === 'success') {
          notify(editingRowIndex ? 'Updated in Google Sheets!' : 'Saved to Google Sheets!', 'success');
          resetForm();
          fetchExpenses(sheetUrl, true);
        } else {
          notify('Google Sheets returned an error while saving.', 'error');
        }
      } catch (err) {
        console.error(err);
        notify('Failed to reach Google Sheets API. Check your script URL and internet connection.', 'error');
      } finally {
        setSubmitting(false);
      }
    } else {
      setTimeout(() => {
        if (editingRowIndex) {
          setExpenses(prev => prev.map(item => item.rowIndex === editingRowIndex ? {
            ...item, date, category: finalCategory, description, amount: numericAmount, paymentMethod, notes, type: txType
          } : item));
          notify('Expense updated in demo mode!', 'success');
        } else {
          const newRow = {
            rowIndex: Date.now(),
            date,
            category: finalCategory,
            description,
            amount: numericAmount,
            type: txType,
            paymentMethod,
            notes,
            createdAt: new Date().toISOString()
          };
          setExpenses(prev => [newRow, ...prev]);
          notify('Saved in local session. Connect Google Sheets for permanent cloud backup!', 'success');
        }
        resetForm();
        setSubmitting(false);
      }, 400);
    }
  };

  const confirmDelete = async () => {
    if (!deleteTargetIndex) return;
    setLoading(true);

    if (isConnected && sheetUrl) {
      try {
        const res = await fetch(sheetUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ action: 'DELETE', rowIndex: deleteTargetIndex })
        });
        const data = await res.json();
        if (data.status === 'success') {
          notify('Transaction deleted from Google Sheets.', 'success');
          fetchExpenses(sheetUrl, true);
        } else {
          notify('Failed to delete row from Google Sheets.', 'error');
        }
      } catch (err) {
        console.error(err);
        notify('Network error attempting to delete row.', 'error');
      } finally {
        setLoading(false);
        setDeleteTargetIndex(null);
      }
    } else {
      setTimeout(() => {
        setExpenses(prev => prev.filter(item => item.rowIndex !== deleteTargetIndex));
        notify('Deleted in demo mode.', 'success');
        setLoading(false);
        setDeleteTargetIndex(null);
      }, 300);
    }
  };

  const startEdit = (item) => {
    setEditingRowIndex(item.rowIndex);
    setDate(item.date);
    setCategory(item.category);
    setDescription(item.description);
    setAmount(item.amount.toString());
    setPaymentMethod(item.paymentMethod);
    setTxType(item.category === 'Income / Salary' || item.type === 'Income' ? 'Income' : 'Expense');
    setNotes(item.notes || '');
    window.scrollTo({ top: 220, behavior: 'smooth' });
  };

  const resetForm = () => {
    setDescription('');
    setAmount('');
    setNotes('');
    setEditingRowIndex(null);
    setTxType('Expense');
    setDate(new Date().toISOString().split('T')[0]);
  };

  const exportToCSV = () => {
    if (expenses.length === 0) {
      notify('No data available to export.', 'error');
      return;
    }
    const headers = ['Date', 'Category', 'Description', 'Amount (INR)', 'Type', 'Payment Method', 'Notes'];
    const rows = expenses.map(e => [
      `"${e.date}"`,
      `"${e.category}"`,
      `"${e.description.replace(/"/g, '""')}"`,
      e.amount,
      `"${e.category === 'Income / Salary' ? 'Income' : 'Expense'}"`,
      `"${e.paymentMethod}"`,
      `"${(e.notes || '').replace(/"/g, '""')}"`
    ]);

    const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `Expense_Report_${new Date().toISOString().split('T')[0]}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    notify('CSV Expense Report downloaded!', 'success');
  };

  const stats = useMemo(() => {
    const expenseItems = expenses.filter(e => e.category !== 'Income / Salary' && e.type !== 'Income');
    const incomeItems = expenses.filter(e => e.category === 'Income / Salary' || e.type === 'Income');

    const totalExpense = expenseItems.reduce((sum, item) => sum + Number(item.amount || 0), 0);
    const totalIncome = incomeItems.reduce((sum, item) => sum + Number(item.amount || 0), 0);
    const count = expenses.length;
    const avgExpense = expenseItems.length > 0 ? totalExpense / expenseItems.length : 0;
    const highestExpense = expenseItems.length > 0 ? Math.max(...expenseItems.map(e => Number(e.amount || 0))) : 0;

    const now = new Date();
    const currentMonthStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const todayStr = now.toISOString().split('T')[0];

    const thisMonthExpense = expenseItems
      .filter(e => (e.date || '').startsWith(currentMonthStr))
      .reduce((sum, e) => sum + Number(e.amount || 0), 0);

    const todayExpense = expenseItems
      .filter(e => e.date === todayStr)
      .reduce((sum, e) => sum + Number(e.amount || 0), 0);

    const budgetPercent = Math.min(Math.round((thisMonthExpense / (monthlyBudget || 1)) * 100), 100);

    return { totalExpense, totalIncome, thisMonthExpense, todayExpense, count, avgExpense, highestExpense, budgetPercent };
  }, [expenses, monthlyBudget]);

  const categoryBreakdown = useMemo(() => {
    const map = {};
    expenses
      .filter(e => e.category !== 'Income / Salary' && e.type !== 'Income')
      .forEach(e => {
        const cat = e.category || 'Other';
        map[cat] = (map[cat] || 0) + Number(e.amount || 0);
      });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [expenses]);

  const filteredExpenses = useMemo(() => {
    return expenses.filter(e => {
      const isInc = e.category === 'Income / Salary' || e.type === 'Income';
      const matchesSearch =
        (e.description || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
        (e.category || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
        (e.notes || '').toLowerCase().includes(searchQuery.toLowerCase());
      const matchesCat = filterCategory === 'All' || e.category === filterCategory;
      const matchesPay = filterPayment === 'All' || e.paymentMethod === filterPayment;
      const matchesType = filterType === 'All' || (filterType === 'Income' ? isInc : !isInc);

      return matchesSearch && matchesCat && matchesPay && matchesType;
    });
  }, [expenses, searchQuery, filterCategory, filterPayment, filterType]);

  const formatINR = (val) => {
    return new Intl.NumberFormat('en-IN', {
      style: 'currency',
      currency: 'INR',
      maximumFractionDigits: 0
    }).format(val || 0);
  };

  return (
    <div className={`min-h-screen font-sans antialiased pb-20 transition-colors duration-200 ${
      darkMode ? 'bg-slate-950 text-slate-100' : 'bg-slate-50/70 text-slate-800'
    }`}>

      {/* Header Bar */}
      <header className={`sticky top-0 z-30 backdrop-blur-md border-b transition-colors ${
        darkMode ? 'bg-slate-900/80 border-slate-800' : 'bg-white/80 border-slate-200/80'
      }`}>
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <span className="p-2.5 bg-gradient-to-tr from-indigo-600 via-indigo-500 to-purple-600 text-white rounded-2xl shadow-lg shadow-indigo-500/25">
              <IndianRupee className="w-6 h-6" />
            </span>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-xl sm:text-2xl font-extrabold tracking-tight">
                  Expense Tracker
                </h1>
                <span className="px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide bg-indigo-100 text-indigo-700 dark:bg-indigo-950 dark:text-indigo-300 rounded-full">
                  Live Sync
                </span>
              </div>
              <p className={`text-xs ${darkMode ? 'text-slate-400' : 'text-slate-500'} font-medium`}>
                Google Sheets-Powered Personal Finance & Budgeting
              </p>
            </div>
          </div>

          <div className="flex items-center flex-wrap gap-2.5">
            <button
              onClick={() => setDarkMode(!darkMode)}
              className={`p-2 rounded-xl border transition-all ${
                darkMode ? 'bg-slate-800 border-slate-700 text-amber-400 hover:bg-slate-700' : 'bg-slate-100 border-slate-200 text-slate-600 hover:bg-slate-200'
              }`}
              title="Toggle Theme"
            >
              {darkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
            </button>

            <button
              onClick={exportToCSV}
              className={`flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-xl border transition-all ${
                darkMode ? 'bg-slate-800 border-slate-700 text-slate-200 hover:bg-slate-700' : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
              }`}
            >
              <Download className="w-3.5 h-3.5" /> CSV Export
            </button>

            {isConnected && (
              <button
                onClick={copyShareLink}
                className="flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-emerald-700 bg-emerald-50 dark:bg-emerald-950/60 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 rounded-xl transition-all"
              >
                <Copy className="w-3.5 h-3.5" />
                {copiedShareLink ? 'Link Copied!' : 'Share Mobile Sync'}
              </button>
            )}

            <button
              onClick={() => setShowCodeModal(true)}
              className="flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-indigo-600 bg-indigo-50 dark:bg-indigo-950/60 dark:text-indigo-300 border border-indigo-200/60 dark:border-indigo-800 rounded-xl transition-all"
            >
              <Code className="w-4 h-4" /> Setup Guide
            </button>

            <button
              onClick={() => setShowPublishModal(true)}
              className="flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold text-white bg-gradient-to-r from-indigo-600 to-purple-600 hover:from-indigo-700 hover:to-purple-700 rounded-xl shadow-md shadow-indigo-500/20 transition-all cursor-pointer"
            >
              <Smartphone className="w-4 h-4" /> Phone Access & Live Hosting
            </button>

            <div className={`flex items-center gap-2 px-3 py-1.5 rounded-full border text-xs font-semibold ${
              isConnected
                ? 'bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-950/40 dark:text-emerald-400 dark:border-emerald-900'
                : 'bg-amber-50 text-amber-700 border-amber-200 dark:bg-amber-950/40 dark:text-amber-400 dark:border-amber-900'
            }`}>
              <span className={`w-2 h-2 rounded-full ${isConnected ? 'bg-emerald-500 animate-pulse' : 'bg-amber-500'}`} />
              {isConnected ? 'Sheets Live' : 'Demo Mode'}
            </div>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-6 space-y-6">

        {/* Notifications */}
        {notification && (
          <div className={`p-4 rounded-2xl border flex items-center gap-3 shadow-lg transition-all animate-in fade-in slide-in-from-top-4 ${
            notification.type === 'success'
              ? 'bg-emerald-600 text-white border-emerald-500'
              : 'bg-rose-600 text-white border-rose-500'
          }`}>
            {notification.type === 'success' ? (
              <CheckCircle2 className="w-5 h-5 shrink-0" />
            ) : (
              <XCircle className="w-5 h-5 shrink-0" />
            )}
            <span className="font-semibold text-sm flex-1">{notification.msg}</span>
            <button onClick={() => setNotification(null)} className="p-1 hover:opacity-80">
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* Google Sheets Sync Connection Box */}
        <section className={`rounded-2xl p-5 border shadow-sm transition-colors ${
          darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
        }`}>
          <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-3 mb-3">
            <div>
              <h2 className="text-sm font-bold flex items-center gap-2">
                <Link2 className="w-4 h-4 text-indigo-500" /> Connect Cloud Database (Google Sheets API)
              </h2>
              <p className={`text-xs mt-0.5 ${darkMode ? 'text-slate-400' : 'text-slate-500'}`}>
                Paste your deployed Apps Script URL to mirror all records in real time across mobile and desktop.
              </p>
            </div>
          </div>

          <form onSubmit={handleConnect} className="flex flex-col sm:flex-row gap-2.5">
            <div className="relative flex-1">
              <input
                type="url"
                placeholder="https://script.google.com/macros/s/AKfycb.../exec"
                value={sheetUrl}
                onChange={(e) => setSheetUrl(e.target.value)}
                className={`w-full pl-3.5 pr-10 py-2 rounded-xl border text-xs font-mono focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                  darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
              />
              {sheetUrl && (
                <button
                  type="button"
                  onClick={() => { setSheetUrl(''); setIsConnected(false); localStorage.removeItem('EXPENSE_TRACKER_SHEET_URL'); }}
                  className="absolute right-3 top-2.5 text-slate-400 hover:text-slate-600"
                >
                  <X className="w-4 h-4" />
                </button>
              )}
            </div>
            <button
              type="submit"
              disabled={loading}
              className="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-semibold transition-all shadow-md shadow-indigo-600/20 disabled:opacity-50 flex items-center justify-center gap-2 shrink-0"
            >
              {loading ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <Link2 className="w-3.5 h-3.5" />}
              {isConnected ? 'Re-Sync' : 'Connect Sheet'}
            </button>
          </form>
        </section>

        {/* Monthly Budget Section */}
        <section className={`p-5 rounded-2xl border shadow-sm ${
          darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
        }`}>
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 mb-3">
            <div className="flex items-center gap-2">
              <Target className="w-5 h-5 text-indigo-500" />
              <div>
                <h3 className="text-sm font-bold">Monthly Spending Budget Goal</h3>
                <p className={`text-xs ${darkMode ? 'text-slate-400' : 'text-slate-500'}`}>
                  Target limit: {formatINR(monthlyBudget)} | Spent: {formatINR(stats.thisMonthExpense)}
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              {isEditingBudget ? (
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    value={monthlyBudget}
                    onChange={(e) => setMonthlyBudget(Number(e.target.value))}
                    className={`w-28 px-2 py-1 text-xs rounded-lg border font-bold ${
                      darkMode ? 'bg-slate-950 border-slate-700' : 'bg-slate-50 border-slate-300'
                    }`}
                  />
                  <button
                    onClick={() => {
                      localStorage.setItem(
                        'EXPENSE_TRACKER_MONTHLY_BUDGET',
                        String(monthlyBudget)
                      );
                      setIsEditingBudget(false);
                   }}
                    className="px-3 py-1 bg-indigo-600 text-white text-xs rounded-lg font-bold"
                  >
                    Save
                  </button>
                </div>
              ) : (
                <button
                  onClick={() => setIsEditingBudget(true)}
                  className={`text-xs font-semibold px-2.5 py-1 rounded-lg border ${
                    darkMode ? 'bg-slate-800 border-slate-700 text-slate-300' : 'bg-slate-100 border-slate-200 text-slate-700'
                  }`}
                >
                  Adjust Goal
                </button>
              )}
            </div>
          </div>

          <div className="w-full bg-slate-200 dark:bg-slate-800 rounded-full h-3 overflow-hidden p-0.5">
            <div
              className={`h-2 rounded-full transition-all duration-500 ${
                stats.budgetPercent > 90 ? 'bg-rose-500' : stats.budgetPercent > 75 ? 'bg-amber-500' : 'bg-emerald-500'
              }`}
              style={{ width: `${stats.budgetPercent}%` }}
            />
          </div>
          <div className="flex justify-between text-[11px] font-semibold mt-1.5 text-slate-400">
            <span>0%</span>
            <span>{stats.budgetPercent}% Used ({formatINR(monthlyBudget - stats.thisMonthExpense)} Remaining)</span>
            <span>100%</span>
          </div>
        </section>

        {/* Dynamic Metric Cards */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <div className={`p-5 rounded-2xl border transition-all ${
            darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
          }`}>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-400">Total Expenses</span>
              <div className="p-2 bg-rose-50 text-rose-600 dark:bg-rose-950 dark:text-rose-400 rounded-lg">
                <TrendingDown className="w-4 h-4" />
              </div>
            </div>
            <h3 className="text-2xl font-black text-rose-600 dark:text-rose-400">{formatINR(stats.totalExpense)}</h3>
            <p className="text-[11px] text-slate-400 mt-1">Cumulative outflow entries</p>
          </div>

          <div className={`p-5 rounded-2xl border transition-all ${
            darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
          }`}>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-400">Total Income</span>
              <div className="p-2 bg-emerald-50 text-emerald-600 dark:bg-emerald-950 dark:text-emerald-400 rounded-lg">
                <TrendingUp className="w-4 h-4" />
              </div>
            </div>
            <h3 className="text-2xl font-black text-emerald-600 dark:text-emerald-400">{formatINR(stats.totalIncome)}</h3>
            <p className="text-[11px] text-slate-400 mt-1">Total credited allowance/salary</p>
          </div>

          <div className={`p-5 rounded-2xl border transition-all ${
            darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
          }`}>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-400">This Month</span>
              <div className="p-2 bg-indigo-50 text-indigo-600 dark:bg-indigo-950 dark:text-indigo-400 rounded-lg">
                <Calendar className="w-4 h-4" />
              </div>
            </div>
            <h3 className="text-2xl font-black text-indigo-600 dark:text-indigo-400">{formatINR(stats.thisMonthExpense)}</h3>
            <p className="text-[11px] text-slate-400 mt-1">Current month expenses</p>
          </div>

          <div className={`p-5 rounded-2xl border transition-all ${
            darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
          }`}>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-400">Net Balance</span>
              <div className="p-2 bg-purple-50 text-purple-600 dark:bg-purple-950 dark:text-purple-400 rounded-lg">
                <IndianRupee className="w-4 h-4" />
              </div>
            </div>
            <h3 className={`text-2xl font-black ${
              stats.totalIncome - stats.totalExpense >= 0 ? 'text-emerald-500' : 'text-rose-500'
            }`}>
              {formatINR(stats.totalIncome - stats.totalExpense)}
            </h3>
            <p className="text-[11px] text-slate-400 mt-1">Savings reserve balance</p>
          </div>
        </div>

        {/* Quick Presets & Smart Parser */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">

          {/* Quick Preset Buttons */}
          <div className={`lg:col-span-1 p-5 rounded-2xl border ${
            darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
          }`}>
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3 flex items-center gap-1.5">
              <Zap className="w-4 h-4 text-amber-500" /> One-Tap Expense Presets
            </h3>
            <div className="flex flex-wrap gap-2">
              {QUICK_PRESETS.map((p, idx) => (
                <button
                  key={idx}
                  onClick={() => handleApplyPreset(p)}
                  className={`px-3 py-1.5 rounded-xl text-xs font-semibold border transition-all ${
                    darkMode
                      ? 'bg-slate-800 border-slate-700 text-slate-200 hover:bg-slate-700 hover:border-indigo-500'
                      : 'bg-slate-50 border-slate-200 text-slate-700 hover:bg-indigo-50 hover:border-indigo-300'
                  }`}
                >
                  {p.label} (₹{p.amount})
                </button>
              ))}
            </div>
          </div>

          {/* AI Smart Text Parser */}
          <div className={`lg:col-span-2 p-5 rounded-2xl border ${
            darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
          }`}>
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3 flex items-center gap-1.5">
              <Wand2 className="w-4 h-4 text-indigo-500" /> Smart Text Expense Parser
            </h3>
            <div className="flex gap-2">
              <input
                type="text"
                placeholder="Type e.g., 'Uber cab to airport for 450 using UPI'..."
                value={aiPrompt}
                onChange={(e) => setAiPrompt(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleAiSmartFill()}
                className={`flex-1 px-3.5 py-2 text-xs rounded-xl border focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                  darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
              />
              <button
                type="button"
                onClick={handleAiSmartFill}
                disabled={isAiProcessing}
                className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-semibold text-xs rounded-xl shadow-md transition-all shrink-0 flex items-center gap-1.5"
              >
                <Sparkles className="w-3.5 h-3.5" /> Parse & Fill
              </button>
            </div>
          </div>
        </div>

        {/* Transaction Entry Form */}
        <section className={`rounded-2xl p-6 border shadow-sm ${
          darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
        }`}>
          <div className="flex items-center justify-between pb-4 border-b border-slate-200/50 mb-6">
            <h2 className="text-base font-bold flex items-center gap-2">
              <PlusCircle className="w-5 h-5 text-indigo-500" />
              {editingRowIndex ? 'Edit Transaction' : 'Record New Transaction'}
            </h2>

            <div className="flex items-center gap-2">
              <div className="flex bg-slate-100 dark:bg-slate-800 p-1 rounded-xl border border-slate-200 dark:border-slate-700">
                <button
                  type="button"
                  onClick={() => setTxType('Expense')}
                  className={`px-3 py-1 rounded-lg text-xs font-bold transition-all ${
                    txType === 'Expense' ? 'bg-rose-600 text-white shadow-xs' : 'text-slate-500'
                  }`}
                >
                  Expense
                </button>
                <button
                  type="button"
                  onClick={() => setTxType('Income')}
                  className={`px-3 py-1 rounded-lg text-xs font-bold transition-all ${
                    txType === 'Income' ? 'bg-emerald-600 text-white shadow-xs' : 'text-slate-500'
                  }`}
                >
                  Income
                </button>
              </div>

              {editingRowIndex && (
                <button
                  type="button"
                  onClick={resetForm}
                  className="text-xs font-semibold text-rose-600 hover:text-rose-700 bg-rose-50 dark:bg-rose-950 px-3 py-1 rounded-lg border border-rose-200 dark:border-rose-800"
                >
                  Cancel Edit
                </button>
              )}
            </div>
          </div>

          <form onSubmit={handleSubmitExpense} className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
            <div>
              <label className="block text-xs font-bold uppercase tracking-wider mb-1.5 text-slate-400">
                Date *
              </label>
              <input
                type="date"
                required
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className={`w-full p-2.5 rounded-xl border text-xs font-medium focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                  darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
              />
            </div>

            <div>
              <label className="block text-xs font-bold uppercase tracking-wider mb-1.5 text-slate-400">
                Category *
              </label>
              <select
                value={category}
                onChange={(e) => setCategory(e.target.value)}
                disabled={txType === 'Income'}
                className={`w-full p-2.5 rounded-xl border text-xs font-medium focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                  darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
              >
                {CATEGORIES.map(cat => (
                  <option key={cat} value={cat}>{cat}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-xs font-bold uppercase tracking-wider mb-1.5 text-slate-400">
                Description *
              </label>
              <input
                type="text"
                required
                placeholder="e.g. Lunch at college canteen"
                value={description}
                onChange={(e) => setDescription(e.target.value)}
                className={`w-full p-2.5 rounded-xl border text-xs font-medium focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                  darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
              />
            </div>

            <div>
              <label className="block text-xs font-bold uppercase tracking-wider mb-1.5 text-slate-400">
                Amount (₹) *
              </label>
              <div className="relative">
                <span className="absolute left-3 top-2.5 text-slate-400 font-bold text-xs">₹</span>
                <input
                  type="number"
                  required
                  min="1"
                  step="any"
                  placeholder="150"
                  value={amount}
                  onChange={(e) => setAmount(e.target.value)}
                  className={`w-full pl-7 pr-3 py-2.5 rounded-xl border text-xs font-bold focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                    darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
                  }`}
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold uppercase tracking-wider mb-1.5 text-slate-400">
                Payment Method *
              </label>
              <select
                value={paymentMethod}
                onChange={(e) => setPaymentMethod(e.target.value)}
                className={`w-full p-2.5 rounded-xl border text-xs font-medium focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                  darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
              >
                {PAYMENT_METHODS.map(pm => (
                  <option key={pm} value={pm}>{pm}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-xs font-bold uppercase tracking-wider mb-1.5 text-slate-400">
                Notes (Optional)
              </label>
              <input
                type="text"
                placeholder="Tags or extra context..."
                value={notes}
                onChange={(e) => setNotes(e.target.value)}
                className={`w-full p-2.5 rounded-xl border text-xs font-medium focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                  darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
              />
            </div>

            <div className="md:col-span-2 lg:col-span-3 pt-2">
              <button
                type="submit"
                disabled={submitting}
                className={`w-full py-3 text-white font-bold text-xs uppercase tracking-wider rounded-xl transition-all shadow-md disabled:opacity-50 flex items-center justify-center gap-2 ${
                  txType === 'Income'
                    ? 'bg-emerald-600 hover:bg-emerald-700 shadow-emerald-600/20'
                    : 'bg-indigo-600 hover:bg-indigo-700 shadow-indigo-600/20'
                }`}
              >
                {submitting ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin" />
                    Saving transaction...
                  </>
                ) : editingRowIndex ? (
                  <>
                    <Check className="w-4 h-4" /> Update Transaction
                  </>
                ) : (
                  <>
                    <PlusCircle className="w-4 h-4" /> Save {txType} Record
                  </>
                )}
              </button>
            </div>
          </form>
        </section>

        {/* Analytics & Architecture Pipeline */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">

          {/* Spending by Category */}
          <div className={`p-6 rounded-2xl border shadow-sm flex flex-col justify-between ${
            darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
          }`}>
            <div>
              <h3 className="text-base font-bold mb-4 flex items-center gap-2">
                <PieChartIcon className="w-5 h-5 text-indigo-500" /> Categorical Spending Breakdown
              </h3>

              {categoryBreakdown.length === 0 ? (
                <div className="text-center py-10 text-slate-400 text-xs">
                  No categorical spending records logged yet.
                </div>
              ) : (
                <div className="space-y-3.5">
                  {categoryBreakdown.map(([cat, total]) => {
                    const pct = Math.round((total / (stats.totalExpense || 1)) * 100);
                    return (
                      <div key={cat} className="space-y-1">
                        <div className="flex justify-between text-xs font-semibold">
                          <span className="flex items-center gap-1.5">
                            <Tag className="w-3.5 h-3.5 text-indigo-500" /> {cat}
                          </span>
                          <span className="font-mono">{formatINR(total)} ({pct}%)</span>
                        </div>
                        <div className="w-full bg-slate-100 dark:bg-slate-800 rounded-full h-2 overflow-hidden">
                          <div
                            className="bg-indigo-600 dark:bg-indigo-500 h-2 rounded-full transition-all duration-500"
                            style={{ width: `${Math.min(pct, 100)}%` }}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>

          {/* Architecture Pipeline Flow */}
          <div className={`p-6 rounded-2xl border shadow-sm flex flex-col justify-between ${
            darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
          }`}>
            <h3 className="text-base font-bold mb-4 flex items-center gap-2">
              <ShieldCheck className="w-5 h-5 text-indigo-500" /> Real-time Cloud Pipeline Architecture
            </h3>

            <div className="py-2 flex flex-col items-center justify-center space-y-2.5">
              {[
                { title: 'Frontend UI', desc: 'React, Tailwind CSS & State validation', color: 'bg-slate-900 text-white' },
                { title: 'Smart AI Parser', desc: 'Natural language text extraction engine', color: 'bg-indigo-600 text-white' },
                { title: 'Google Apps Script', desc: 'Serverless REST endpoint webhook', color: 'bg-blue-600 text-white' },
                { title: 'Google Sheet Database', desc: 'Cloud multi-device persistent storage', color: 'bg-emerald-600 text-white' }
              ].map((step, idx) => (
                <React.Fragment key={step.title}>
                  <div className={`w-full p-3 rounded-xl ${step.color} shadow-xs flex items-center justify-between`}>
                    <span className="text-xs font-bold uppercase tracking-wider">{step.title}</span>
                    <span className="text-[11px] opacity-80 font-normal">{step.desc}</span>
                  </div>
                  {idx < 3 && (
                    <ArrowRight className="w-4 h-4 text-slate-400 rotate-90 my-0.5" />
                  )}
                </React.Fragment>
              ))}
            </div>
          </div>
        </div>

        {/* Expense History Table */}
        <section className={`rounded-2xl p-6 border shadow-sm ${
          darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'
        }`}>
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
            <div>
              <h2 className="text-lg font-bold">Transaction History</h2>
              <p className="text-xs text-slate-400">Showing newest entries first</p>
            </div>
            <button
              onClick={() => fetchExpenses(sheetUrl, false)}
              disabled={loading}
              className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-xl border transition-all ${
                darkMode ? 'bg-slate-800 border-slate-700 hover:bg-slate-700' : 'bg-slate-50 border-slate-200 hover:bg-slate-100'
              }`}
            >
              <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin text-indigo-500' : ''}`} />
              Refresh
            </button>
          </div>

          {/* Filter Bar */}
          <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 mb-6">
            <div className="relative">
              <Search className="w-4 h-4 absolute left-3 top-2.5 text-slate-400" />
              <input
                type="text"
                placeholder="Search notes, item..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className={`w-full pl-9 pr-3 py-2 rounded-xl border text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                  darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
                }`}
              />
            </div>

            <select
              value={filterType}
              onChange={(e) => setFilterType(e.target.value)}
              className={`py-2 px-3 border rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
              }`}
            >
              <option value="All">All Types (Income & Expense)</option>
              <option value="Expense">Expenses Only</option>
              <option value="Income">Income Only</option>
            </select>

            <select
              value={filterCategory}
              onChange={(e) => setFilterCategory(e.target.value)}
              className={`py-2 px-3 border rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
              }`}
            >
              <option value="All">All Categories</option>
              {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
            </select>

            <select
              value={filterPayment}
              onChange={(e) => setFilterPayment(e.target.value)}
              className={`py-2 px-3 border rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500/20 ${
                darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'
              }`}
            >
              <option value="All">All Payment Methods</option>
              {PAYMENT_METHODS.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
          </div>

          {/* Table */}
          {loading ? (
            <div className="text-center py-12 text-slate-400 text-xs">
              <RefreshCw className="w-6 h-6 animate-spin mx-auto text-indigo-500 mb-2" />
              Loading records...
            </div>
          ) : filteredExpenses.length === 0 ? (
            <div className="text-center py-12 border-2 border-dashed border-slate-300 dark:border-slate-800 rounded-2xl">
              <p className="font-bold text-sm">No transactions found.</p>
              <p className="text-slate-400 text-xs mt-1">Add your first expense or clear search filters.</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-slate-200 dark:border-slate-800 text-[11px] font-bold text-slate-400 uppercase tracking-wider">
                    <th className="py-3 px-4">Date</th>
                    <th className="py-3 px-4">Category</th>
                    <th className="py-3 px-4">Description</th>
                    <th className="py-3 px-4">Amount</th>
                    <th className="py-3 px-4">Payment</th>
                    <th className="py-3 px-4">Notes</th>
                    <th className="py-3 px-4 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-800/60 text-xs font-medium">
                  {filteredExpenses.map((item) => {
                    const isIncome = item.category === 'Income / Salary' || item.type === 'Income';
                    return (
                      <tr key={item.rowIndex || item.createdAt} className="hover:bg-slate-50/50 dark:hover:bg-slate-800/40 transition-colors">
                         <td className="py-3.5 px-4 font-mono text-slate-400 whitespace-nowrap">{item.date.split('-')[2] + ' ' + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][Number(item.date.split('-')[1]) - 1] + ' ' + item.date.split('-')[0]}</td>
                             <td className="py-3.5 px-4 whitespace-nowrap">     
                               <span className={`px-2.5 py-1 rounded-full font-bold text-[10px] tracking-wide uppercase border ${
                            isIncome
                              ? 'bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-950 dark:text-emerald-300 dark:border-emerald-800'
                              : 'bg-slate-100 text-slate-800 border-slate-200 dark:bg-slate-800 dark:text-slate-200 dark:border-slate-700'
                          }`}>
                            {item.category}
                          </span>
                        </td>
                        <td className="py-3.5 px-4 font-semibold">{item.description}</td>
                        <td className={`py-3.5 px-4 font-black whitespace-nowrap ${
                          isIncome ? 'text-emerald-500' : 'text-rose-500'
                        }`}>
                          {isIncome ? '+' : '-'}{formatINR(item.amount)}
                        </td>
                        <td className="py-3.5 px-4 text-slate-400 whitespace-nowrap">{item.paymentMethod}</td>
                        <td className="py-3.5 px-4 text-slate-400 italic max-w-xs truncate">{item.notes || '-'}</td>
                        <td className="py-3.5 px-4 text-right whitespace-nowrap space-x-1">
                          <button
                            onClick={() => startEdit(item)}
                            className="p-1.5 text-slate-400 hover:text-indigo-500 rounded-lg hover:bg-indigo-50 dark:hover:bg-indigo-950 transition-colors"
                            title="Edit"
                          >
                            <Edit3 className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => setDeleteTargetIndex(item.rowIndex)}
                            className="p-1.5 text-slate-400 hover:text-rose-500 rounded-lg hover:bg-rose-50 dark:hover:bg-rose-950 transition-colors"
                            title="Delete"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </section>

      </main>

      {/* Delete Confirmation Modal */}
      {deleteTargetIndex && (
        <div className="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className={`rounded-2xl p-6 max-w-sm w-full shadow-2xl border ${
            darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-100'
          }`}>
            <h3 className="text-base font-bold">Confirm Transaction Deletion</h3>
            <p className="text-xs text-slate-400 mt-2">
              Are you sure you want to delete this row? This action will permanently remove it from your spreadsheet.
            </p>
            <div className="flex items-center gap-3 mt-6">
              <button
                onClick={() => setDeleteTargetIndex(null)}
                className={`flex-1 py-2 text-xs font-semibold rounded-xl ${
                  darkMode ? 'bg-slate-800 text-slate-300' : 'bg-slate-100 text-slate-600'
                }`}
              >
                Cancel
              </button>
              <button
                onClick={confirmDelete}
                className="flex-1 py-2 text-xs font-semibold text-white bg-rose-600 hover:bg-rose-700 rounded-xl shadow-md shadow-rose-600/20"
              >
                Yes, Delete
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Code / Setup Guide Modal */}
      {showCodeModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className={`rounded-2xl p-6 max-w-2xl w-full shadow-2xl border my-8 ${
            darkMode ? 'bg-slate-900 border-slate-800 text-slate-200' : 'bg-white border-slate-100 text-slate-700'
          }`}>
            <div className="flex items-center justify-between pb-3 border-b border-slate-200/40">
              <h3 className="text-base font-bold flex items-center gap-2">
                <Code className="w-5 h-5 text-indigo-500" /> Connect Your Google Sheet in 1 Minute
              </h3>
              <button
                onClick={() => setShowCodeModal(false)}
                className="p-1 text-slate-400 hover:text-slate-600"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-4 my-4 text-xs">
              <ol className="list-decimal pl-4 space-y-2 leading-relaxed">
                <li>Create a blank document in <b>Google Sheets</b>.</li>
                <li>In menu bar select: <b>Extensions</b> &gt; <b>Apps Script</b>.</li>
                <li>Paste the script below into the code editor:</li>
              </ol>

              <div className="relative">
                <pre className="p-4 bg-slate-950 text-indigo-300 rounded-xl text-[11px] font-mono overflow-x-auto max-h-60 leading-relaxed border border-slate-800">
                  {APPS_SCRIPT_CODE}
                </pre>
                <button
                  onClick={async () => {
                    const success = await safeCopyToClipboard(APPS_SCRIPT_CODE);
                    if (success) {
                      setCopiedCode(true);
                      setTimeout(() => setCopiedCode(false), 2500);
                    } else {
                      notify('Could not copy code automatically.', 'error');
                    }
                  }}
                  className="absolute right-3 top-3 px-3 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-[11px] font-semibold flex items-center gap-1.5 shadow-md"
                >
                  {copiedCode ? <Check className="w-3 h-3" /> : <Copy className="w-3 h-3" />}
                  {copiedCode ? 'Copied!' : 'Copy Script'}
                </button>
              </div>

              <ol start="4" className="list-decimal pl-4 space-y-2 leading-relaxed">
                <li>Click <b>Deploy</b> &gt; <b>New deployment</b>.</li>
                <li>Select type: <b>Web app</b>.</li>
                <li>Set <b>Execute as</b>: <i>Me</i>.</li>
                <li>Set <b>Who has access</b>: <i>Anyone</i> (CRITICAL for app access).</li>
                <li>Click <b>Deploy</b>, copy the <b>Web App URL</b>, and paste it in the Connect box above!</li>
              </ol>
            </div>

            <button
              onClick={() => setShowCodeModal(false)}
              className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-xl shadow-md"
            >
              Done / Return to App
            </button>
          </div>
        </div>
      )}

      {/* Publish Modal */}
      {showPublishModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className={`rounded-2xl p-6 max-w-2xl w-full shadow-2xl border my-8 ${
            darkMode ? 'bg-slate-900 border-slate-800 text-slate-200' : 'bg-white border-slate-100 text-slate-700'
          }`}>
            <div className="flex items-center justify-between pb-3 border-b border-slate-200/40">
              <h3 className="text-base font-bold flex items-center gap-2">
                <Smartphone className="w-5 h-5 text-indigo-500" /> Publish Live & Mobile Phone Access
              </h3>
              <button
                onClick={() => setShowPublishModal(false)}
                className="p-1 text-slate-400 hover:text-slate-600"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-5 my-4 text-xs">
              <div className="p-4 rounded-xl bg-indigo-50 dark:bg-indigo-950/50 border border-indigo-200 dark:border-indigo-800 flex flex-col sm:flex-row items-center gap-4">
                <div className="bg-white p-2 rounded-xl shadow-md shrink-0 border border-slate-200">
                  <img
                    src={`https://api.qrserver.com/v1/create-qr-code/?size=140x140&data=${encodeURIComponent(
                      typeof window !== 'undefined' ? window.location.href : 'https://vercel.com'
                    )}`}
                    alt="Scan QR Code to open on mobile"
                    className="w-32 h-32"
                  />
                </div>
                <div className="space-y-1.5 text-center sm:text-left">
                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wide bg-indigo-600 text-white">
                    <QrCode className="w-3 h-3" /> Scan with Phone Camera
                  </span>
                  <h4 className="font-bold text-sm text-slate-900 dark:text-white">Instant Mobile Sync</h4>
                  <p className="text-slate-600 dark:text-slate-300 text-xs">
                    Scan this QR code with your iPhone or Android camera to open this live web app directly on your smartphone.
                  </p>
                  <p className="text-[11px] text-indigo-600 dark:text-indigo-400 font-semibold pt-1">
                    💡 Tip: Tap "Add to Home Screen" on your phone to use it like a native mobile app icon!
                  </p>
                </div>
              </div>

              <div className="space-y-3">
                <h4 className="font-bold text-sm flex items-center gap-1.5 text-slate-900 dark:text-white">
                  <Globe className="w-4 h-4 text-emerald-500" /> How to Publish Live Online for FREE
                </h4>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className={`p-3.5 rounded-xl border ${darkMode ? 'bg-slate-950 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
                    <div className="font-bold text-indigo-600 dark:text-indigo-400 flex items-center justify-between mb-1">
                      <span>Method 1: Vercel or Netlify</span>
                      <ExternalLink className="w-3.5 h-3.5" />
                    </div>
                    <ol className="list-decimal pl-4 space-y-1 text-[11px] leading-relaxed text-slate-600 dark:text-slate-300">
                      <li>Export or download your React code files.</li>
                      <li>Go to <b>Vercel.com</b> or <b>Netlify.com</b> (100% Free).</li>
                      <li>Connect your GitHub repo or drag & drop the project folder.</li>
                      <li>Get your live <code>.vercel.app</code> website link in 1 minute!</li>
                    </ol>
                  </div>

                  <div className={`p-3.5 rounded-xl border ${darkMode ? 'bg-slate-950 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
                    <div className="font-bold text-emerald-600 dark:text-emerald-400 flex items-center justify-between mb-1">
                      <span>Method 2: GitHub Pages</span>
                      <ExternalLink className="w-3.5 h-3.5" />
                    </div>
                    <ol className="list-decimal pl-4 space-y-1 text-[11px] leading-relaxed text-slate-600 dark:text-slate-300">
                      <li>Push this project code into a <b>GitHub repository</b>.</li>
                      <li>Go to <b>Settings</b> &gt; <b>Pages</b>.</li>
                      <li>Select <b>Main Branch</b> and click <b>Save</b>.</li>
                      <li>Your website will be live at <code>yourname.github.io</code>!</li>
                    </ol>
                  </div>
                </div>
              </div>

              <div className={`p-3.5 rounded-xl border ${darkMode ? 'bg-slate-950 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
                <h5 className="font-bold mb-1 text-slate-800 dark:text-slate-200 flex items-center gap-1.5">
                  <Share2 className="w-3.5 h-3.5 text-purple-500" /> Save as Phone App Icon
                </h5>
                <ul className="list-disc pl-4 space-y-1 text-[11px] text-slate-600 dark:text-slate-300">
                  <li><b>iPhone (Safari):</b> Open your live URL &rarr; Tap <i>Share</i> button (square with arrow) &rarr; Tap <b>"Add to Home Screen"</b>.</li>
                  <li><b>Android (Chrome):</b> Open your live URL &rarr; Tap <i>3 Dots</i> (top right) &rarr; Tap <b>"Add to Home Screen"</b> / <b>"Install App"</b>.</li>
                </ul>
              </div>
            </div>

            <button
              onClick={() => setShowPublishModal(false)}
              className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-xl shadow-md cursor-pointer"
            >
              Close Guide
            </button>
          </div>
        </div>
      )}

    </div>
  );
}
