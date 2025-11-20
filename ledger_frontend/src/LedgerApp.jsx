import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { List, Plus, Minus, CreditCard, Save, Trash2, X, Clipboard, ArrowDownUp, ArrowUpDown, Edit, ChevronUp, ChevronDown, FileDown, Loader2, FileUp, BookOpen } from 'lucide-react';

// --- Ledger API & Utility Setup ---
const DEFAULT_APP_ID = typeof __app_id !== 'undefined' ? __app_id : 'ledger-app';
const DEFAULT_LEDGER_API_BASE = typeof __ledger_api_base !== 'undefined' ? __ledger_api_base : '/api/ledger';
const LEDGER_TOKEN_STORAGE_KEY = 'ledger_session_token';
const PENDING_IMPORT_STORAGE_KEY = 'pending_ledger_imports';

const appId = DEFAULT_APP_ID;
const ledgerApiBase = DEFAULT_LEDGER_API_BASE;

const getStoredLedgerToken = () => {
    if (typeof window === 'undefined' || !window.localStorage) {
        return null;
    }
    try {
        return window.localStorage.getItem(LEDGER_TOKEN_STORAGE_KEY);
    } catch (error) {
        console.warn('Failed to read ledger token from storage:', error);
        return null;
    }
};

const persistLedgerToken = (token) => {
    if (typeof window === 'undefined' || !window.localStorage) {
        return;
    }
    try {
        window.localStorage.setItem(LEDGER_TOKEN_STORAGE_KEY, token);
    } catch (error) {
        console.warn('Failed to store ledger token:', error);
    }
};

const loadPendingImportsFromStorage = () => {
    if (typeof window === 'undefined' || !window.localStorage) {
        return [];
    }
    try {
        const raw = window.localStorage.getItem(PENDING_IMPORT_STORAGE_KEY);
        if (!raw) {
            return [];
        }
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed : [];
    } catch (error) {
        console.warn('Failed to read pending ledger imports:', error);
        return [];
    }
};

const savePendingImportsToStorage = (entries) => {
    if (typeof window === 'undefined' || !window.localStorage) {
        return;
    }
    try {
        window.localStorage.setItem(PENDING_IMPORT_STORAGE_KEY, JSON.stringify(entries));
    } catch (error) {
        console.warn('Failed to store pending ledger imports:', error);
    }
};

// データ構造の定義
// Account: { id: string, name: string, number: string, order: number, userId: string }
// Transaction: { id: string, accountId: string, date: string, withdrawal: number, deposit: number, memo: string, type: string, timestamp: object, userOrder?: number }

// 通貨フォーマット（日本円を想定）
const formatCurrency = (value) => {
  if (value === undefined || value === null || isNaN(value)) return '';
  return new Intl.NumberFormat('ja-JP', { style: 'decimal' }).format(value);
};

// JSONから読み込んだデータ内のFirestoreタイムスタンプ形式をJSのDateオブジェクトに再帰的に変換するヘルパー関数
const convertFirestoreTimestamps = (data) => {
    if (data === null || typeof data !== 'object') {
        return data;
    }

    // Firestoreのタイムスタンプ形式（秒・ナノ秒）をチェック
    if (typeof data.seconds === 'number' && typeof data.nanoseconds === 'number') {
        return new Date(data.seconds * 1000);
    }
    
    // 配列の場合は各要素を再帰的に処理
    if (Array.isArray(data)) {
        return data.map(item => convertFirestoreTimestamps(item));
    }

    // オブジェクトの場合は各プロパティを再帰的に処理
    const newData = {};
    for (const key in data) {
        newData[key] = convertFirestoreTimestamps(data[key]);
    }
    return newData;
};

const baseOrderValue = (transaction) => {
    if (!transaction) return 0;
    if (typeof transaction.userOrder === 'number') {
        return transaction.userOrder;
    }
    const dateValue = new Date(transaction.date).getTime();
    return Number.isFinite(dateValue) ? dateValue : 0;
};

const resolveHolderName = (account) => {
    if (!account) return '';
    if (account.holderName) return account.holderName;
    if (account.holder_name) return account.holder_name;
    if (account.holder) return account.holder;
    if (Array.isArray(account.ownerName)) {
        return account.ownerName.filter(Boolean).join(' / ');
    }
    if (Array.isArray(account.owner_name)) {
        return account.owner_name.filter(Boolean).join(' / ');
    }
    return '';
};

const sortTransactionsByConfig = (transactions, sort, accountMap = {}) => {
    if (!Array.isArray(transactions)) return [];
    if (!sort || sort.field === 'custom') {
        return [...transactions];
    }
    const direction = sort.direction === 'desc' ? -1 : 1;

    const getComparableValue = (transaction) => {
        switch (sort.field) {
            case 'date': {
                const timestamp = new Date(transaction.date).getTime();
                return Number.isFinite(timestamp) ? timestamp : 0;
            }
            case 'withdrawal':
                return Number(transaction.withdrawal) || 0;
            case 'deposit':
                return Number(transaction.deposit) || 0;
            case 'memo':
                return (transaction.memo || '').toString().toLowerCase();
            case 'account':
                return (accountMap[transaction.accountId]?.name || '').toLowerCase();
            case 'holder':
                return resolveHolderName(accountMap[transaction.accountId]).toLowerCase();
            default:
                return baseOrderValue(transaction);
        }
    };

    return [...transactions].sort((a, b) => {
        const valueA = getComparableValue(a);
        const valueB = getComparableValue(b);

        if (typeof valueA === 'string' && typeof valueB === 'string') {
            const comparison = valueA.localeCompare(valueB, 'ja');
            if (comparison !== 0) {
                return comparison * direction;
            }
        } else if (valueA !== valueB) {
            return (valueA > valueB ? 1 : -1) * direction;
        }

        const fallback = baseOrderValue(a) - baseOrderValue(b);
        if (fallback !== 0) {
            return fallback;
        }
        return a.id.localeCompare(b.id);
    });
};

const generateClientId = () => {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return crypto.randomUUID();
    }
    return `client_${Date.now().toString(36)}_${Math.random().toString(36).slice(2)}`;
};

const normalizeAccountForImport = (account) => ({
    id: account.id,
    name: account.name,
    number: account.number,
    holderName: account.holder_name || account.holderName || '',
    order: typeof account.order === 'number' ? account.order : 0,
});

const normalizeTransactionForImport = (transaction) => ({
    id: transaction.id,
    accountId: transaction.accountId || transaction.account_id,
    date: transaction.date,
    withdrawal: transaction.withdrawal ?? transaction.withdrawal_amount ?? 0,
    deposit: transaction.deposit ?? transaction.deposit_amount ?? 0,
    memo: transaction.memo || '',
    type: transaction.type || '',
    userOrder: transaction.userOrder ?? transaction.user_order ?? null,
    rowColor: transaction.rowColor ?? transaction.row_color ?? null,
});

const buildMergedImportPayload = ({
    caseId,
    existingAccounts,
    existingTransactions,
    incomingAccounts,
    incomingTransactions,
}) => {
    const mergedAccounts = (existingAccounts || []).map(normalizeAccountForImport);
    const mergedTransactions = (existingTransactions || []).map(normalizeTransactionForImport);
    const existingAccountIds = new Set(mergedAccounts.map((account) => account.id));
    let maxOrder = mergedAccounts.reduce((max, account) => Math.max(max, account.order || 0), 0);
    const accountIdMap = new Map();

    (incomingAccounts || []).forEach((incoming, index) => {
        const sourceId = incoming.id || generateClientId();
        let accountId = sourceId;
        if (existingAccountIds.has(accountId)) {
            accountId = generateClientId();
        }
        existingAccountIds.add(accountId);
        accountIdMap.set(sourceId, accountId);
        mergedAccounts.push({
            id: accountId,
            name: incoming.name,
            number: incoming.number,
            holderName: incoming.holderName || incoming.holder_name || '',
            order: maxOrder + (index + 1) * 1000,
        });
    });

    (incomingTransactions || []).forEach((incomingTxn) => {
        const mappedAccountId = accountIdMap.get(incomingTxn.accountId || incomingTxn.account_id) || (incomingTxn.accountId || incomingTxn.account_id);
        mergedTransactions.push({
            id: incomingTxn.id || generateClientId(),
            accountId: mappedAccountId,
            date: incomingTxn.date,
            withdrawal: incomingTxn.withdrawal ?? incomingTxn.withdrawal_amount ?? 0,
            deposit: incomingTxn.deposit ?? incomingTxn.deposit_amount ?? 0,
            memo: incomingTxn.memo || incomingTxn.description || '',
            type: incomingTxn.type || incomingTxn.description || '',
            userOrder: incomingTxn.userOrder ?? incomingTxn.user_order ?? null,
            rowColor: incomingTxn.rowColor ?? incomingTxn.row_color ?? null,
        });
    });

    return {
        caseId,
        accounts: mergedAccounts,
        transactions: mergedTransactions,
    };
};

const INITIAL_TRANSACTION_FILTER = {
    accountId: 'all',
    direction: 'all',
    keyword: '',
    minAmount: '',
    maxAmount: '',
    rowColor: 'all',
};

const INSURANCE_KEYWORDS = [
    '保険',
    '生命',
    'ひまわり生命',
    '第一生命',
    '日本生命',
    '明治安田',
    '住友生命',
    'ソニー生命',
    'かんぽ生命',
    'アフラック',
    'JA共済',
    '共済',
];

const isLikelyPersonalMemo = (memo) => {
    if (!memo) return false;
    const normalized = memo.trim();
    if (!normalized) return false;
    if (/様|さま|さん/.test(normalized)) return true;
    if (/[(（]/.test(normalized)) return false;
    const lettersOnly = normalized.replace(/[0-9\s]/g, '');
    if (!lettersOnly) return false;
    const pattern = /^[\p{sc=Han}\p{sc=Hiragana}\p{sc=Katakana}A-Za-z]{2,}$/u;
    return pattern.test(lettersOnly) && lettersOnly.length <= 12;
};

const runComplianceAnalysis = (accounts, transactions) => {
    const result = { generatedAt: new Date().toISOString(), checks: [] };
    const insuranceAccounts = (accounts || []).filter((account) => {
        const target = `${account.name || ''}${account.holder_name || account.holderName || ''}`;
        return INSURANCE_KEYWORDS.some((keyword) => target.includes(keyword));
    });
    if (insuranceAccounts.length > 0) {
        result.checks.push({
            category: '保険契約の有無',
            severity: 'info',
            items: insuranceAccounts.map((account) => `${account.name || account.number || account.id}`),
            message: `${insuranceAccounts.length}件の口座に保険関連の名称が含まれています。解約返戻金や契約者貸付の有無をご確認ください。`,
        });
    }

    const giftMap = new Map();
    (transactions || []).forEach((txn) => {
        if (!txn?.deposit) return;
        if (!txn?.date) return;
        const memo = (txn.memo || txn.type || '').trim();
        if (!isLikelyPersonalMemo(memo)) return;
        const year = new Date(txn.date).getFullYear();
        if (!giftMap.has(memo)) {
            giftMap.set(memo, new Map());
        }
        const yearMap = giftMap.get(memo);
        yearMap.set(year, (yearMap.get(year) || 0) + (txn.deposit || 0));
    });
    const giftFindings = [];
    giftMap.forEach((yearMap, memo) => {
        yearMap.forEach((amount, year) => {
            if (amount >= 1100000) {
                giftFindings.push({ memo, year, amount });
            }
        });
    });
    if (giftFindings.length > 0) {
        result.checks.push({
            category: '贈与税の検討',
            severity: 'warn',
            items: giftFindings.map((item) => `${item.year}年 ${item.memo}: ${formatCurrency(item.amount)}円`),
            message: '個人名義への入金で年間110万円を超えるものがあります。贈与税申告の有無を確認してください。',
        });
    }

    if (result.checks.length === 0) {
        result.checks.push({
            category: '特記事項',
            severity: 'info',
            items: [],
            message: '該当する注意項目は検出されませんでした。',
        });
    }
    return result;
};


// --- Custom Components ---

// ローディングとエラーメッセージの表示
const StatusMessage = ({ loading, error, userId }) => (
  <div className="p-4 bg-white/90 backdrop-blur-sm shadow-xl rounded-xl">
    {loading && <p className="text-blue-600 font-semibold flex items-center justify-center">データをロード中...</p>}
    {error && <p className="text-red-600 font-semibold">エラーが発生しました: {error.message}</p>}
    {userId && <p className="text-xs text-gray-500 mt-2 text-center">ユーザーID (データ保存先): <span className="font-mono text-xs">{userId}</span></p>}
  </div>
);

// モーダルコンポーネント
const Modal = ({ children, title, isOpen, onClose, className = 'max-w-lg' }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 bg-gray-900 bg-opacity-75 flex items-center justify-center z-50 p-4 overflow-y-auto">
      <div className={`bg-white rounded-xl shadow-2xl w-full ${className} transform transition-all duration-300 scale-100 opacity-100 flex flex-col max-h-[90vh]`}>
        <div className="flex justify-between items-center p-4 border-b flex-shrink-0">
          <h3 className="text-xl font-bold text-gray-800">{title}</h3>
          <button onClick={onClose} className="text-gray-500 hover:text-gray-700 p-1 rounded-full hover:bg-gray-100 transition">
            <X size={20} />
          </button>
        </div>
        <div className="p-6 overflow-y-auto">
          {children}
        </div>
      </div>
    </div>
  );
};

// 入力フォームコンポーネント
const InputField = ({ label, id, type = 'text', value, onChange, placeholder = '', required = false, icon: Icon, className = '', inputMode = 'text', pattern = null }) => (
  <div className={`mb-4 ${className}`}>
    <label htmlFor={id} className="block text-sm font-medium text-gray-700 mb-1">{label}</label>
    <div className="relative">
      {Icon && <Icon className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" size={18} />}
      <input
        id={id}
        type={type}
        value={value}
        onChange={onChange}
        placeholder={placeholder}
        required={required}
        inputMode={inputMode}
        pattern={pattern}
        className={`w-full p-2.5 border border-gray-300 rounded-lg focus:ring-blue-500 focus:border-blue-500 transition duration-150 shadow-sm ${Icon ? 'pl-10' : ''}`}
      />
    </div>
  </div>
);

// 選択フィールドコンポーネント
const SelectField = ({ label, id, value, onChange, options, required = false, className = '' }) => (
    <div className={`mb-4 ${className}`}>
        <label htmlFor={id} className="block text-sm font-medium text-gray-700 mb-1">{label}</label>
        <select
            id={id}
            value={value}
            onChange={onChange}
            required={required}
            className="w-full p-2.5 border border-gray-300 rounded-lg focus:ring-blue-500 focus:border-blue-500 transition duration-150 shadow-sm bg-white"
        >
            {options.map(option => (
                <option key={option.value} value={option.value} disabled={option.disabled}>
                    {option.label}
                </option>
            ))}
        </select>
    </div>
);


// 金額入力コンポーネント
const CurrencyInput = ({ label, id, value, onChange, isDeposit = false }) => {
  const handleChange = (e) => {
    // 数字とカンマ以外の文字を削除
    const rawValue = e.target.value.replace(/[^0-9,]/g, '');
    const numberValue = parseInt(rawValue.replace(/,/g, ''), 10);
    onChange(isNaN(numberValue) ? '' : numberValue);
  };

  return (
    <div className="mb-4">
      <label htmlFor={id} className="block text-sm font-medium text-gray-700 mb-1">{label}</label>
      <div className="relative">
        {isDeposit ? 
          <Plus className="absolute left-3 top-1/2 transform -translate-y-1/2 text-green-500" size={18} /> : 
          <Minus className="absolute left-3 top-1/2 transform -translate-y-1/2 text-red-500" size={18} />
        }
        <input
          id={id}
          type="text"
          inputMode="numeric"
          pattern="[0-9,]*"
          value={formatCurrency(value)}
          onChange={handleChange}
          placeholder="0"
          className={`w-full p-2.5 border border-gray-300 rounded-lg focus:ring-blue-500 focus:border-blue-500 transition duration-150 shadow-sm pl-10 text-right font-mono ${isDeposit ? 'text-green-700' : 'text-red-700'}`}
        />
      </div>
    </div>
  );
};

// メインボタンコンポーネント
const MainButton = ({ children, onClick, Icon, className = 'bg-blue-600 hover:bg-blue-700', disabled = false, type = 'button' }) => (
  <button
    onClick={onClick}
    disabled={disabled}
    type={type}
    className={`flex items-center justify-center space-x-2 px-6 py-3 text-white font-semibold rounded-xl shadow-lg transition duration-200 transform hover:scale-[1.02] active:scale-[0.98] focus:outline-none focus:ring-4 focus:ring-opacity-50 ${className} disabled:bg-gray-400 disabled:shadow-none disabled:transform-none`}
  >
    {Icon && <Icon size={20} />}
    <span>{children}</span>
  </button>
);

// --- PDF Export Button ---
const ExportPDFButton = ({ elementId, fileName, title = 'PDFに出力' }) => {
    const [exportStatus, setExportStatus] = useState('idle'); // 'idle', 'loading-libs', 'ready', 'exporting', 'error'
    const [errorMessage, setErrorMessage] = useState('');

    useEffect(() => {
        // コンポーネントのマウント時にライブラリの読み込みを開始
        if (window.jspdf && window.html2canvas) {
            setExportStatus('ready');
            return;
        }

        setExportStatus('loading-libs');

        const loadScript = (src) => new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = src;
            script.async = true;
            script.onload = resolve;
            script.onerror = () => reject(new Error(`${src}の読み込みに失敗しました。`));
            document.body.appendChild(script);
        });

        Promise.all([
            loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"),
            loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js")
        ]).then(() => {
            setExportStatus('ready');
        }).catch(error => {
            console.error(error);
            setErrorMessage('PDFライブラリの読み込みに失敗しました。');
            setExportStatus('error');
        });
    }, []);

    const handleExport = async () => {
        const input = document.getElementById(elementId);
        if (!input) {
            setErrorMessage('PDF出力対象の要素が見つかりません。');
            setExportStatus('error');
            return;
        }

        setExportStatus('exporting');
        setErrorMessage('');

        // Find the table body to count the rows
        const tableBody = input.querySelector('tbody');
        const allRows = tableBody ? Array.from(tableBody.querySelectorAll('tr')) : [];
        const totalRows = allRows.length;

        try {
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF({
                orientation: 'p',
                unit: 'pt',
                format: 'a4'
            });

            // If there are no rows, just print the current view as a single page
            if (totalRows === 0) {
                const canvas = await window.html2canvas(input, {
                    scale: 2,
                    useCORS: true,
                    backgroundColor: '#ffffff'
                });
                const imgData = canvas.toDataURL('image/png');
                const pdfWidth = pdf.internal.pageSize.getWidth();
                const imgProps = pdf.getImageProperties(imgData);
                const ratio = pdfWidth / imgProps.width;
                const imgHeight = imgProps.height * ratio;
                pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, imgHeight);

            } else {
                // Logic for paginated export based on row count
                const rowsPerPageFirst = 14;
                const rowsPerPageNext = 17;
                let rowsProcessed = 0;
                let pageNum = 0;

                while (rowsProcessed < totalRows) {
                    pageNum++;
                    const isFirstPage = (pageNum === 1);
                    const rowsThisPage = isFirstPage ? rowsPerPageFirst : rowsPerPageNext;
                    
                    const startIndex = rowsProcessed;
                    const endIndex = Math.min(startIndex + rowsThisPage, totalRows);
                    const isLastPage = endIndex === totalRows;

                    // Create a clone of the original container for rendering each page
                    const pageClone = input.cloneNode(true);
                    document.body.appendChild(pageClone); // Append to body to compute styles

                    // --- PDF EXPORT: Hide unnecessary columns ---
                    const tableInClone = pageClone.querySelector('table');
                    if (tableInClone) {
                        // Hide header columns ('着色', '操作')
                        const headerRow = tableInClone.querySelector('thead tr');
                        if (headerRow) {
                            if (headerRow.children[6]) headerRow.children[6].style.display = 'none'; // 着色
                            if (headerRow.children[7]) headerRow.children[7].style.display = 'none'; // 順序
                            if (headerRow.children[8]) headerRow.children[8].style.display = 'none'; // 操作
                        }

                        // Hide body columns ('着色', '編集', '削除')
                        const bodyRows = tableInClone.querySelectorAll('tbody tr');
                        bodyRows.forEach(row => {
                            if (row.children.length > 9) { // Make sure it is a data row
                                if (row.children[6]) row.children[6].style.display = 'none';
                                if (row.children[7]) row.children[7].style.display = 'none';
                                if (row.children[8]) row.children[8].style.display = 'none';
                                if (row.children[9]) row.children[9].style.display = 'none';
                            }
                        });

                        // Adjust footer colspan for the balance cell
                        const footerRow = tableInClone.querySelector('tfoot tr');
                        if (footerRow && footerRow.children.length === 4) {
                            const balanceCell = footerRow.children[3];
                            balanceCell.setAttribute('colSpan', '1');
                        }
                    }
                    // --- END PDF EXPORT MODIFICATION ---


                    // Isolate the rows for the current page
                    const cloneBody = pageClone.querySelector('tbody');
                    const cloneRows = Array.from(cloneBody.querySelectorAll('tr'));

                    cloneRows.forEach((row, index) => {
                        // Keep row only if it's in the current page's range
                        if (index < startIndex || index >= endIndex) {
                           row.style.display = 'none';
                        }
                    });

                    // Hide footer if it's not the last page
                    const cloneFooter = pageClone.querySelector('tfoot');
                    if (cloneFooter && !isLastPage) {
                        cloneFooter.style.display = 'none';
                    }
                    
                    // Prepare clone for rendering off-screen
                    pageClone.style.position = 'absolute';
                    pageClone.style.left = '0px';
                    pageClone.style.top = '0px';
                    pageClone.style.zIndex = '-1'; // Render below other content
                    pageClone.style.width = `${input.offsetWidth}px`;
                    
                    const canvas = await html2canvas(pageClone, {
                        scale: 2,
                        useCORS: true,
                        backgroundColor: '#ffffff',
                    });

                    document.body.removeChild(pageClone);

                    // Add the rendered canvas to the PDF
                    const imgData = canvas.toDataURL('image/png');
                    const pdfWidth = pdf.internal.pageSize.getWidth();
                    const pdfHeight = pdf.internal.pageSize.getHeight();
                    
                    const imgProps = pdf.getImageProperties(imgData);
                    const ratio = pdfWidth / imgProps.width;
                    const imgHeight = imgProps.height * ratio;

                    // Ensure the image fits within the page height
                    const finalHeight = imgHeight > pdfHeight ? pdfHeight : imgHeight;

                    if (!isFirstPage) {
                        pdf.addPage();
                    }
                    pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, finalHeight);

                    rowsProcessed = endIndex;
                }
            }

            pdf.save(`${fileName}.pdf`);
            setExportStatus('ready');

        } catch (error) {
            console.error("Error exporting to PDF:", error);
            setErrorMessage('PDFの生成中にエラーが発生しました。');
            setExportStatus('error');
        }
    };
    
    const renderButtonContent = () => {
        switch (exportStatus) {
            case 'loading-libs':
                return <><Loader2 size={16} className="animate-spin" /> <span>ライブラリ準備中...</span></>;
            case 'exporting':
                return <><Loader2 size={16} className="animate-spin" /> <span>PDF生成中...</span></>;
            case 'error':
                return <span>エラー</span>;
            default:
                return <><FileDown size={16} /> <span>{title}</span></>;
        }
    };

    return (
        <div className="flex flex-col items-end">
            <button
                onClick={handleExport}
                disabled={exportStatus !== 'ready'}
                className="flex items-center justify-center space-x-2 px-4 py-2 text-white font-semibold rounded-lg shadow-md transition duration-200 transform hover:scale-[1.02] active:scale-[0.98] focus:outline-none focus:ring-2 focus:ring-offset-2 bg-teal-600 hover:bg-teal-700 focus:ring-teal-500 disabled:bg-gray-400 disabled:shadow-none disabled:transform-none disabled:cursor-not-allowed text-sm"
            >
                {renderButtonContent()}
            </button>
            {exportStatus === 'error' && <p className="text-xs text-red-600 mt-1">{errorMessage}</p>}
        </div>
    );
};


// --- Edit Transaction Modal Component ---

const EditTransactionModal = ({ isOpen, onClose, transaction, onUpdateTransaction }) => {
    const [date, setDate] = useState('');
    const [withdrawal, setWithdrawal] = useState('');
    const [deposit, setDeposit] = useState('');
    const [memo, setMemo] = useState('');
    const [type, setType] = useState('振込'); // 新しい取引種別
    const [message, setMessage] = useState('');

    useEffect(() => {
        if (transaction) {
            setDate(transaction.date || '');
            setWithdrawal(transaction.withdrawal > 0 ? transaction.withdrawal : '');
            setDeposit(transaction.deposit > 0 ? transaction.deposit : '');
            setMemo(transaction.memo || '');
            setType(transaction.type || '振込');
            setMessage('');
        }
    }, [transaction]);

    if (!transaction) return null;

    const transactionTypes = [
        { value: '振込', label: '振込 (銀行)' },
        { value: '振替', label: '振替 (口座間)' },
        { value: '現金', label: '現金' }
    ];

    const handleUpdateTransaction = async (e) => {
        e.preventDefault();
        
        const withdrawalValue = parseInt(withdrawal, 10) || 0;
        const depositValue = parseInt(deposit, 10) || 0;

        if (!date || (withdrawalValue === 0 && depositValue === 0)) {
            setMessage('日付と金額を入力してください。');
            return;
        }

        if (withdrawalValue > 0 && depositValue > 0) {
            setMessage('出金と入金は同時に登録できません。');
            return;
        }

        try {
            await onUpdateTransaction(transaction.id, {
                date,
                withdrawal: withdrawalValue,
                deposit: depositValue,
                memo,
                type,
            });
            setMessage('取引情報が更新されました！');
            setTimeout(onClose, 1500);
        } catch (e) {
            console.error("Error updating transaction:", e);
            setMessage(`更新エラー: ${e.message}`);
        }
    };

    return (
        <Modal isOpen={isOpen} title="取引情報の編集" onClose={onClose} className="max-w-xl">
            <form onSubmit={handleUpdateTransaction}>
                <p className="mb-4 text-sm text-gray-600">口座名: <span className="font-semibold">{transaction.accountName}</span></p>
                
                <div className="grid grid-cols-2 gap-4">
                    <InputField
                        label="日付"
                        id="editDate"
                        type="date"
                        value={date}
                        onChange={(e) => setDate(e.target.value)}
                        required
                    />
                    <SelectField
                        label="取引種別"
                        id="editType"
                        value={type}
                        onChange={(e) => setType(e.target.value)}
                        options={transactionTypes}
                        required
                    />
                </div>
                {pendingImports.length > 0 && (
                    <div className="mb-4 text-sm bg-amber-50 border border-amber-200 rounded-xl p-3 flex flex-wrap items-center justify-between gap-3">
                        <span className="text-amber-900">未登録の通帳データが {pendingImports.length} 件あります。案件に取り込みましょう。</span>
                        <MainButton onClick={() => setShowPendingImportModal(true)} className="bg-amber-600 hover:bg-amber-700">
                            取り込みを開始
                        </MainButton>
                    </div>
                )}
                
                <CurrencyInput 
                    label="出金額"
                    id="editWithdrawal"
                    value={withdrawal}
                    onChange={setWithdrawal}
                />
                <CurrencyInput
                    label="入金額"
                    id="editDeposit"
                    value={deposit}
                    onChange={setDeposit}
                    isDeposit={true}
                />
                
                <InputField
                    label="備考"
                    id="editMemo"
                    value={memo}
                    onChange={(e) => setMemo(e.target.value)}
                    placeholder="取引内容を簡単にメモ"
                />

                {message && <p className={`p-3 rounded-lg my-3 text-sm ${message.includes('エラー') ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}`}>{message}</p>}

                <MainButton type="submit" Icon={Save} className="w-full mt-4 bg-green-600 hover:bg-green-700">
                    変更を保存
                </MainButton>
            </form>
        </Modal>
    );
};


// --- Add Account Modal Component (New) ---

const AddAccountModal = ({ isOpen, onClose, onCreateAccount, caseName }) => {
  const [holder, setHolder] = useState('');
  const [name, setName] = useState('');
  const [number, setNumber] = useState('');
  const [message, setMessage] = useState('');

    const handleSaveAccount = async () => {
        if (!holder || !name || !number) {
            setMessage('名義人・口座名・口座番号を入力してください。');
            return;
        }

        try {
            await onCreateAccount({ name, number, holderName: holder });
            setMessage('口座情報が正常に登録されました！');
            setHolder('');
            setName('');
            setNumber('');
            setTimeout(onClose, 1500);
        } catch (e) {
            console.error("Error adding account:", e);
            setMessage(`登録エラー: ${e.message}`);
        }
    };

  return (
    <Modal isOpen={isOpen} title="新規取引口座の登録" onClose={onClose}>
      {caseName && <p className="text-sm text-gray-500 mb-3">案件: {caseName}</p>}
      <form onSubmit={(e) => { e.preventDefault(); handleSaveAccount(); }}>
        <InputField label="名義人" id="accountHolder" value={holder} onChange={(e) => setHolder(e.target.value)} required icon={Clipboard} placeholder="山田 太郎" />
        <InputField label="口座表示名" id="accountName" value={name} onChange={(e) => setName(e.target.value)} required icon={List} placeholder="生活費 口座" />
        <InputField label="口座番号" id="accountNumber" value={number} onChange={(e) => setNumber(e.target.value.replace(/[^0-9]/g, ''))} required icon={CreditCard} placeholder="1234567" type="tel" inputMode="numeric" pattern="[0-9]*" />
        
        {message && <p className={`p-3 rounded-lg my-3 ${message.includes('エラー') ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}`}>{message}</p>}

        <MainButton type="submit" Icon={Plus} className="w-full mt-4 bg-green-600 hover:bg-green-700">
          新規口座を登録
        </MainButton>
      </form>
    </Modal>
  );
};

const EditAccountModal = ({ isOpen, onClose, account, onUpdateAccount }) => {
    const [holder, setHolder] = useState(account?.holder_name || account?.holderName || '');
    const [name, setName] = useState(account?.name || '');
    const [number, setNumber] = useState(account?.number || '');
    const [message, setMessage] = useState('');

    useEffect(() => {
        setHolder(account?.holder_name || account?.holderName || '');
        setName(account?.name || '');
        setNumber(account?.number || '');
        setMessage('');
    }, [account, isOpen]);

    if (!account) {
        return null;
    }

    const handleSubmit = async (event) => {
        event.preventDefault();
        try {
            await onUpdateAccount(account.id, {
                name,
                number,
                holderName: holder,
            });
            setMessage('口座情報を更新しました。');
            setTimeout(() => {
                setMessage('');
                onClose();
            }, 1200);
        } catch (error) {
            console.error('Failed to update account:', error);
            setMessage(error.message || '更新に失敗しました。');
        }
    };

    return (
        <Modal isOpen={isOpen} onClose={onClose} title="口座情報を編集">
            <form onSubmit={handleSubmit} className="space-y-4">
                <InputField
                    label="名義人"
                    id="editHolder"
                    value={holder}
                    onChange={(e) => setHolder(e.target.value)}
                    icon={Clipboard}
                />
                <InputField
                    label="口座表示名"
                    id="editAccountName"
                    value={name}
                    onChange={(e) => setName(e.target.value)}
                    icon={List}
                />
                <InputField
                    label="口座番号"
                    id="editAccountNumber"
                    value={number}
                    onChange={(e) => setNumber(e.target.value.replace(/[^0-9]/g, ''))}
                    icon={CreditCard}
                    inputMode="numeric"
                    pattern="[0-9]*"
                />
                {message && (
                    <p className={`p-3 rounded-lg my-2 text-sm ${message.includes('失敗') ? 'bg-red-50 text-red-600' : 'bg-green-50 text-green-700'}`}>
                        {message}
                    </p>
                )}
                <div className="flex justify-end gap-3">
                    <button type="button" className="text-sm text-gray-500" onClick={onClose}>キャンセル</button>
                    <MainButton type="submit" Icon={Save} className="bg-blue-600 hover:bg-blue-700">
                        更新する
                    </MainButton>
                </div>
            </form>
        </Modal>
    );
};

const UsageGuideModal = ({ isOpen, onClose }) => {
    return (
        <Modal isOpen={isOpen} onClose={onClose} title="入出金検討表の使い方" className="max-w-3xl">
            <div className="space-y-5 text-sm text-gray-700 leading-relaxed">
                <p className="text-base text-gray-800">
                    SOROBOCRで読み取った <code className="bg-gray-100 text-gray-800 px-1 rounded">bank_transactions.json</code> をそのまま案件へ取り込み、通帳単位で色付け・ソートできます。
                    3分で把握できる流れをまとめました。
                </p>
                <ol className="list-decimal list-inside space-y-3">
                    <li>
                        <span className="font-semibold">案件を作成/選択</span>…上部の案件セレクターで対象案件を決め、必要なら「案件を追加」で新規作成します。
                    </li>
                    <li>
                        <span className="font-semibold">OCR結果を自動連携</span>…SOROBOCRで読み取り後に表示される「入出金検討表ツールを開く」から遷移すると、ブラウザに保存済みの通帳候補が検出されます。黄色のカードで取り込み方法を選択してください。
                    </li>
                    <li>
                        <span className="font-semibold">口座情報を整える</span>…名義と口座表示名をそれぞれ入力し、口座順を整えます。案件ごとにFastAPI (Railway) 上へ保存され、ブラウザを閉じても復元できます。
                    </li>
                    <li>
                        <span className="font-semibold">統合タブで検討</span>…統合タブではすべての取引をリスト化。今回追加したソートバーで日付や金額順に並び替え、最終的な優先度は「手動順序」で調整してください。
                    </li>
                    <li>
                        <span className="font-semibold">PDF/JSONで共有</span>…統合タブ右上のPDFボタンや、設定タブのエクスポートからJSONを出力し、レビュー資料に貼り付けられます。
                    </li>
                </ol>
                <div className="bg-indigo-50 border border-indigo-100 text-indigo-900 rounded-lg p-4 space-y-2">
                    <p className="text-sm font-semibold">さらに詳しい手順</p>
                    <p>より詳細な画面遷移やFAQは別タブのガイドページにまとめています。</p>
                    <a
                        href="./guide.html"
                        target="_blank"
                        rel="noopener"
                        className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-indigo-600 text-white text-sm font-semibold hover:bg-indigo-700"
                    >
                        <BookOpen size={16} /> ガイドページを開く
                    </a>
                </div>
            </div>
        </Modal>
    );
};


// --- Data Import/Export Modals ---

const ExportModal = ({ isOpen, onClose, accounts, transactions }) => {
    const [fileName, setFileName] = useState(`取引データ_${new Date().toISOString().split('T')[0]}`);
    const [message, setMessage] = useState('');

    const handleExport = () => {
        if (!fileName) {
            setMessage('ファイル名を入力してください。');
            return;
        }

        setMessage('');
        const dataToExport = {
            accounts: accounts,
            transactions: transactions,
            exportedAt: new Date().toISOString(),
        };

        const jsonString = JSON.stringify(dataToExport, null, 2);
        const blob = new Blob([jsonString], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${fileName}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        onClose();
    };
    
    return (
        <Modal isOpen={isOpen} title="データのエクスポート" onClose={onClose}>
            <p className="text-sm text-gray-600 mb-4">
                すべての口座情報と取引履歴を一つのJSONファイルに書き出します。
            </p>
            <InputField 
                label="ファイル名 (.json)"
                id="exportFileName"
                value={fileName}
                onChange={(e) => setFileName(e.target.value)}
                placeholder="例: backup_2025-10-12"
            />
            {message && <p className="text-red-500 text-sm mt-2">{message}</p>}
            <MainButton onClick={handleExport} Icon={FileDown} className="w-full mt-4 bg-teal-600 hover:bg-teal-700">
                エクスポート実行
            </MainButton>
        </Modal>
    );
};

const ImportModal = ({ isOpen, onClose, onImport, caseName }) => {
    const [file, setFile] = useState(null);
    const [isConfirmed, setIsConfirmed] = useState(false);
    const [status, setStatus] = useState('idle'); // 'idle', 'processing', 'success', 'error'
    const [message, setMessage] = useState('');
    
    const handleFileChange = (e) => {
        setFile(e.target.files[0]);
        setMessage('');
        setStatus('idle');
    };

    const handleImport = async () => {
        if (!file) {
            setMessage('インポートするJSONファイルを選択してください。');
            return;
        }
        if (!isConfirmed) {
            setMessage('上書きの確認チェックボックスをオンにしてください。');
            return;
        }

        setStatus('processing');
        setMessage('インポート処理を実行中... この処理には時間がかかる場合があります。');
        
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const data = JSON.parse(e.target.result);
                if (!data.accounts || !data.transactions) {
                    throw new Error('ファイルの形式が正しくありません。(accounts or transactions not found)');
                }
                setMessage('Railway APIへデータを送信しています...');
                const accountsToImport = data.accounts.map(convertFirestoreTimestamps);
                const transactionsToImport = data.transactions.map(convertFirestoreTimestamps);
                await onImport({ accounts: accountsToImport, transactions: transactionsToImport });
                setStatus('success');
                setMessage('データのインポートが完了しました。');
                setTimeout(() => {
                    onClose();
                }, 2000);

            } catch (error) {
                console.error("Import error:", error);
                setStatus('error');
                setMessage(`インポートに失敗しました: ${error.message}`);
            }
        };
        reader.readAsText(file);
    };

    useEffect(() => {
        // モーダルが閉じたときに状態をリセット
        if (!isOpen) {
            setFile(null);
            setIsConfirmed(false);
            setStatus('idle');
            setMessage('');
        }
    }, [isOpen]);

    return (
        <Modal isOpen={isOpen} title="データのインポート" onClose={onClose}>
            <div className="bg-red-50 border-l-4 border-red-500 text-red-700 p-4 mb-4 rounded" role="alert">
                <p className="font-bold">【重要】データの復元に関する警告</p>
                <p className="text-sm">
                    この操作を行うと、{caseName ? `${caseName} に` : ''}現在登録されている<span className="font-semibold">すべての口座と取引履歴が完全に削除</span>され、
                    ファイルの内容に置き換わります。この操作は元に戻せません。
                </p>
            </div>
            
            <label className="block mb-2 text-sm font-medium text-gray-900" htmlFor="file_input">バックアップファイル(.json)を選択</label>
            <input 
                className="block w-full text-sm text-gray-900 border border-gray-300 rounded-lg cursor-pointer bg-gray-50 focus:outline-none" 
                id="file_input" 
                type="file"
                accept=".json"
                onChange={handleFileChange}
                disabled={status === 'processing'}
            />

            <div className="mt-4 flex items-start">
                <div className="flex items-center h-5">
                    <input 
                        id="confirm-overwrite" 
                        type="checkbox" 
                        checked={isConfirmed}
                        onChange={(e) => setIsConfirmed(e.target.checked)}
                        disabled={status === 'processing'}
                        className="w-4 h-4 text-blue-600 bg-gray-100 border-gray-300 rounded focus:ring-blue-500"
                    />
                </div>
                <div className="ml-3 text-sm">
                    <label htmlFor="confirm-overwrite" className="font-medium text-gray-900">
                        現在の全データが削除され上書きされることを理解し、インポートを続行します。
                    </label>
                </div>
            </div>

            {message && (
                <p className={`p-3 rounded-lg my-3 text-sm ${
                    status === 'error' ? 'bg-red-100 text-red-700' : 
                    status === 'success' ? 'bg-green-100 text-green-700' :
                    'bg-blue-100 text-blue-700'
                }`}>{message}</p>
            )}

            <MainButton 
                onClick={handleImport} 
                Icon={status === 'processing' ? Loader2 : FileUp} 
                className="w-full mt-4 bg-red-600 hover:bg-red-700"
                disabled={!file || !isConfirmed || status === 'processing'}
            >
                {status === 'processing' ? '処理中...' : 'インポート実行'}
            </MainButton>
        </Modal>
    );
};

const PendingImportModal = ({
    isOpen,
    onClose,
    pendingImports,
    caseName,
    onApply,
    onManual,
    onDismiss,
    status,
    error,
}) => {
    const [newCaseNames, setNewCaseNames] = useState({});
    if (!isOpen) return null;
    return (
        <Modal isOpen={isOpen} title="未登録の口座候補" onClose={onClose} className="max-w-3xl">
            <p className="text-sm text-gray-600 mb-4">
                {caseName ? `${caseName} に` : '選択中の案件に'} 取り込める通帳データがブラウザに保存されています。自動取り込みを実行すると、口座と取引が案件へ追加されます。
            </p>
            <div className="space-y-4 max-h-[50vh] overflow-y-auto">
                {pendingImports.length === 0 && <p className="text-sm text-gray-500">未登録のデータはありません。</p>}
                {pendingImports.map((entry) => (
                    <div key={entry.id} className="border border-blue-200 rounded-xl p-4 bg-blue-50">
                        <div className="flex justify-between items-center mb-2">
                            <div>
                                <h4 className="text-lg font-semibold text-blue-900">{entry.name || '未命名の通帳'}</h4>
                                <p className="text-xs text-blue-800">保存日時: {new Date(entry.createdAt).toLocaleString()}</p>
                            </div>
                            <div className="flex gap-2">
                                <button
                                    type="button"
                                    className="px-3 py-1 text-xs rounded-full bg-white border border-red-200 text-red-600"
                                    onClick={() => onDismiss(entry.id)}
                                >
                                    破棄
                                </button>
                            </div>
                        </div>
                        <div className="space-y-3">
                            <p className="text-sm text-blue-900">口座候補: {(entry.assets || []).length} 件</p>
                            <MainButton
                                onClick={() => onApply(entry)}
                                Icon={status === 'applying' ? Loader2 : Save}
                                className="bg-blue-600 hover:bg-blue-700"
                                disabled={status === 'applying'}
                            >
                                {status === 'applying' ? '取り込み中…' : '選択中の案件に取り込む'}
                            </MainButton>
                            <div className="bg-white border border-blue-200 rounded-lg p-3 space-y-2">
                                <p className="text-xs text-blue-900">新しい案件として取り込む</p>
                                <input
                                    type="text"
                                    value={newCaseNames[entry.id] || ''}
                                    onChange={(e) => setNewCaseNames((prev) => ({ ...prev, [entry.id]: e.target.value }))}
                                    placeholder={`${entry.name || '通帳'}の案件名`}
                                    className="w-full p-2 border rounded-md text-sm"
                                />
                                <MainButton
                                    onClick={() => onApply(entry, { newCaseName: newCaseNames[entry.id] || entry.name || '新規案件' })}
                                    Icon={status === 'applying' ? Loader2 : Save}
                                    className="bg-amber-600 hover:bg-amber-700"
                                    disabled={status === 'applying'}
                                >
                                    {status === 'applying' ? '取り込み中…' : '新しい案件に登録'}
                                </MainButton>
                            </div>
                        </div>
                    </div>
                ))}
            </div>
            {error && <p className="text-sm text-red-600 bg-red-100 rounded-lg p-2 mt-3">{error}</p>}
            <div className="flex justify-end gap-3 mt-4">
                <button className="text-sm text-gray-600 underline" onClick={onManual}>手動で口座を登録する</button>
                <MainButton onClick={onClose} className="bg-gray-600 hover:bg-gray-700">閉じる</MainButton>
            </div>
        </Modal>
    );
};

// --- Account Management Content (For the 'register' tab) ---

const AccountManagementContent = ({
    accounts,
    caseName,
    setShowExportModal,
    setShowImportModal,
    onReorderAccountOrder,
    onDeleteAccount,
    onAddAccountClick,
    onEditAccount,
}) => {
    const [message, setMessage] = useState('');

    // 口座の順序を変更する関数
    const handleReorderAccount = async (currentAccount, targetAccount) => {
        if (currentAccount.id === targetAccount.id) return;
        try {
            await onReorderAccountOrder([
                { id: currentAccount.id, order: targetAccount.order },
                { id: targetAccount.id, order: currentAccount.order },
            ]);
            setMessage('口座の順序を変更しました。');
        } catch (e) {
            console.error('Error reordering account:', e);
            setMessage(`順序変更エラー: ${e.message}`);
        }
        setTimeout(() => setMessage(''), 1500);
    };

    const handleDeleteAccount = async (accountId) => {
        // カスタムモーダルを使う代わりに、一時的にwindow.confirmを使用
        if (!window.confirm('この口座と、関連するすべての取引を削除してもよろしいですか？')) return;

        try {
            await onDeleteAccount(accountId);
            setMessage('口座を削除しました。');
        } catch (e) {
            console.error('Error deleting account:', e);
            setMessage(`削除エラー: ${e.message}`);
        }
        setTimeout(() => setMessage(''), 3000);
    };


    // 順序でソートされた口座リスト
    const sortedAccounts = useMemo(() => {
        return [...accounts].sort((a, b) => (a.order || 9999) - (b.order || 9999));
    }, [accounts]);

    return (
        <div className="p-6 space-y-6">
            <h2 className="text-2xl font-bold text-gray-800 border-b pb-2 flex justify-between items-center">
                <span>{caseName ? `${caseName} の口座管理` : '取引口座の登録・管理と並べ替え'}</span>
                <MainButton Icon={Plus} onClick={onAddAccountClick} className="bg-green-600 hover:bg-green-700 px-4 py-2 text-base">
                    新規口座を登録
                </MainButton>
            </h2>

            <p className="text-gray-600">
                登録されている口座を一覧表示し、ドラッグ＆ドロップの代わりに上下ボタンで表示順序を変更できます。
            </p>

            {message && <p className={`p-3 rounded-lg my-3 text-sm ${message.includes('エラー') ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}`}>{message}</p>}

            {/* データ管理ボタン */}
            <div className="bg-gray-100 p-4 rounded-lg flex items-center justify-start space-x-4 border border-gray-200">
                <h3 className="text-lg font-semibold text-gray-700">データ管理:</h3>
                <MainButton Icon={FileDown} onClick={() => setShowExportModal(true)} className="bg-teal-600 hover:bg-teal-700 px-4 py-2 text-sm">
                    エクスポート
                </MainButton>
                <MainButton Icon={FileUp} onClick={() => setShowImportModal(true)} className="bg-red-600 hover:bg-red-700 px-4 py-2 text-sm">
                    インポート
                </MainButton>
            </div>


            <div className="pt-2">
                <h4 className="text-lg font-semibold mb-3 text-gray-700">登録済み口座 ({accounts.length}件)</h4>
                <div className="space-y-3 p-2 bg-gray-50 rounded-xl shadow-inner max-h-[50vh] overflow-y-auto">
                {sortedAccounts.map((acc, index) => (
                    <div key={acc.id} className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between bg-white p-4 rounded-lg border border-gray-200 shadow-sm">
                    <div className="text-sm space-y-2 w-full">
                        <div className="grid gap-3 sm:grid-cols-2">
                            <div>
                                <p className="text-xs font-semibold text-gray-500 tracking-wide uppercase">名義人</p>
                                <p className="text-base text-gray-900">{resolveHolderName(acc) || '未登録'}</p>
                            </div>
                            <div>
                                <p className="text-xs font-semibold text-gray-500 tracking-wide uppercase">口座表示名</p>
                                <p className="text-lg font-bold text-gray-900">{acc.name || '---'}</p>
                            </div>
                        </div>
                        <div className="flex flex-wrap gap-2 text-xs text-gray-500">
                            <span className="inline-flex items-center px-2 py-0.5 rounded-full bg-gray-100 font-mono">No. {acc.number || '---'}</span>
                            <span className="inline-flex items-center px-2 py-0.5 rounded-full bg-gray-100">表示順: {index + 1}番目</span>
                        </div>
                    </div>
                    <div className="flex space-x-3 items-center">
                        {/* 順序変更ボタン */}
                        <div className="flex flex-col space-y-0.5 border-r pr-3">
                            {index > 0 && (
                                <button onClick={() => handleReorderAccount(acc, sortedAccounts[index - 1])} title="上へ移動" className="text-gray-500 hover:text-blue-600 p-0.5 rounded-full hover:bg-blue-50 transition">
                                    <ChevronUp size={20} />
                                </button>
                            )}
                            {index < sortedAccounts.length - 1 && (
                                <button onClick={() => handleReorderAccount(acc, sortedAccounts[index + 1])} title="下へ移動" className="text-gray-500 hover:text-blue-600 p-0.5 rounded-full hover:bg-blue-50 transition">
                                    <ChevronDown size={20} />
                                </button>
                            )}
                        </div>
                        <button
                            onClick={() => onEditAccount(acc)}
                            className="text-blue-500 hover:text-blue-700 p-2 rounded-full hover:bg-blue-50 transition"
                            title="口座情報を編集"
                        >
                            <Edit size={18} />
                        </button>
                        <button
                            onClick={() => handleDeleteAccount(acc.id)}
                            className="text-red-500 hover:text-red-700 p-2 rounded-full hover:bg-red-100 transition"
                            title="この口座を削除"
                        >
                            <Trash2 size={20} />
                        </button>
                    </div>
                    </div>
                ))}
                {accounts.length === 0 && <p className="text-sm text-gray-500 text-center py-6">まだ口座が登録されていません。上のボタンから登録してください。</p>}
                </div>
            </div>
        </div>
    );
};


// --- Transaction Input & List Component (Single Account Tab Content) ---

const TransactionTabContent = ({ account, transactions, onCreateTransaction, onDeleteTransaction, setEditingTransaction }) => {
    const [date, setDate] = useState(new Date().toISOString().substring(0, 10));
    const [withdrawal, setWithdrawal] = useState('');
    const [deposit, setDeposit] = useState('');
    const [memo, setMemo] = useState('');
    const [type, setType] = useState('振込'); // 新しい取引種別
    const [message, setMessage] = useState('');
    const [localSort, setLocalSort] = useState({ field: 'date', direction: 'asc' });

    const accountTransactions = useMemo(() => {
        return transactions.filter(t => t.accountId === account.id);
    }, [transactions, account.id]);

    const accountMap = useMemo(() => ({ [account.id]: account }), [account]);
    const sortedTransactions = useMemo(() => sortTransactionsByConfig(accountTransactions, localSort, accountMap), [accountTransactions, localSort, accountMap]);

    const transactionTypes = [
        { value: '振込', label: '振込 (銀行)' },
        { value: '振替', label: '振替 (口座間)' },
        { value: '現金', label: '現金' }
    ];

    const handleSaveTransaction = async (e) => {
        e.preventDefault();
        
        const withdrawalValue = parseInt(withdrawal, 10) || 0;
        const depositValue = parseInt(deposit, 10) || 0;

        if (!date || (withdrawalValue === 0 && depositValue === 0)) {
            setMessage('日付と、出金額または入金額の少なくとも一方を入力してください。');
            return;
        }

        if (withdrawalValue > 0 && depositValue > 0) {
            setMessage('出金と入金は同時に登録できません。どちらか一方のみを入力してください。');
            return;
        }

        try {
            await onCreateTransaction({
                accountId: account.id,
                date,
                withdrawal: withdrawalValue,
                deposit: depositValue,
                memo,
                type,
            });
            setMessage('取引を登録しました！');
            setWithdrawal('');
            setDeposit('');
            setMemo('');
            setType('振込');
        } catch (e) {
            console.error('Error adding transaction:', e);
            setMessage(`取引登録エラー: ${e.message}`);
        }
        setTimeout(() => setMessage(''), 3000);
    };

    const handleDeleteTransaction = async (transactionId) => {
        if (!window.confirm('この取引を削除してもよろしいですか？')) return;

        try {
            await onDeleteTransaction(transactionId);
            setMessage('取引を削除しました。');
        } catch (e) {
            console.error('Error deleting transaction:', e);
            setMessage(`取引削除エラー: ${e.message}`);
        }
        setTimeout(() => setMessage(''), 3000);
    };

    return (
        <div className="p-4 space-y-6">
            <h2 className="text-2xl font-bold text-gray-800 border-b pb-2">
                <span className="text-blue-600">{account.name}</span> の取引入力
            </h2>

            {/* 取引入力フォーム */}
            <form onSubmit={handleSaveTransaction} className="bg-white p-6 rounded-xl shadow-md border border-gray-100">
                <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
                    <InputField
                        label="日付"
                        id="transactionDate"
                        type="date"
                        value={date}
                        onChange={(e) => setDate(e.target.value)}
                        required
                        className="md:col-span-1"
                    />
                    <SelectField
                        label="種別"
                        id="transactionType"
                        value={type}
                        onChange={(e) => setType(e.target.value)}
                        options={transactionTypes}
                    />
                    <CurrencyInput 
                        label="出金額"
                        id="withdrawal"
                        value={withdrawal}
                        onChange={setWithdrawal}
                    />
                    <CurrencyInput
                        label="入金額"
                        id="deposit"
                        value={deposit}
                        onChange={setDeposit}
                        isDeposit={true}
                    />
                    <InputField
                        label="備考"
                        id="memo"
                        value={memo}
                        onChange={(e) => setMemo(e.target.value)}
                        placeholder="取引内容を簡単にメモ"
                    />
                </div>
                
                <div className="flex justify-end items-center mt-4">
                    <MainButton type="submit" Icon={Save} className="w-1/3 bg-blue-600 hover:bg-blue-700">
                        この取引を登録
                    </MainButton>
                </div>
                
                {message && <p className={`p-3 rounded-lg my-3 text-sm ${message.includes('エラー') ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}`}>{message}</p>}

            </form>

            {/* 取引履歴リスト */}
            <div className="mt-6">
                <div className="flex flex-wrap items-start justify-between gap-3 mb-3">
                    <div>
                        <h3 className="text-xl font-bold text-gray-700">取引履歴 ({accountTransactions.length}件)</h3>
                        <p className="text-sm text-gray-500">日付や金額で並べ替えてから編集できます。</p>
                    </div>
                    <div className="flex flex-wrap items-center gap-2 text-sm">
                        <label className="text-gray-500">並び替え</label>
                        <select
                            value={localSort.field}
                            onChange={(e) => {
                                const value = e.target.value;
                                setLocalSort((prev) => ({
                                    field: value,
                                    direction: value === prev.field ? prev.direction : (value === 'withdrawal' || value === 'deposit' ? 'desc' : 'asc'),
                                }));
                            }}
                            className="border border-gray-300 rounded-lg px-3 py-1.5 bg-white"
                        >
                            <option value="date">日付</option>
                            <option value="withdrawal">出金額</option>
                            <option value="deposit">入金額</option>
                            <option value="memo">摘要</option>
                        </select>
                        <button
                            type="button"
                            onClick={() => setLocalSort((prev) => ({ field: prev.field, direction: prev.direction === 'asc' ? 'desc' : 'asc' }))}
                            className="inline-flex items-center gap-1 px-3 py-1.5 border border-gray-300 rounded-lg text-gray-700 hover:bg-gray-50"
                        >
                            <ArrowUpDown size={16} /> {localSort.direction === 'asc' ? '昇順' : '降順'}
                        </button>
                    </div>
                </div>
                <div className="bg-white rounded-xl shadow-lg overflow-hidden">
                    <TransactionTable
                        transactions={sortedTransactions}
                        accounts={[{id: account.id, name: account.name}]} // Pass the current account in an array
                        onDelete={handleDeleteTransaction}
                        onEdit={setEditingTransaction}
                        onColorChange={() => {}} // Dummy function as color is not changed here
                        showAccountInfo={false} // We are in a single account view
                        sorting={localSort}
                        onSortField={(field) => {
                            setLocalSort((prev) => {
                                if (prev.field === field) {
                                    return { field, direction: prev.direction === 'asc' ? 'desc' : 'asc' };
                                }
                                return { field, direction: field === 'withdrawal' || field === 'deposit' ? 'desc' : 'asc' };
                            });
                        }}
                        sortableFields={['date', 'withdrawal', 'deposit', 'memo']}
                    />
                </div>
            </div>
        </div>
    );
};

const TransactionTable = ({ transactions, accounts, onDelete, onEdit, onColorChange, onReorder, showAccountInfo = true, sorting = null, onSortField = null, sortableFields = [], highlightAccountIds = [] }) => {
    const accountMap = useMemo(() => {
        if (!accounts) return {};
        return accounts.reduce((map, acc) => {
            map[acc.id] = acc;
            return map;
        }, {});
    }, [accounts]);

    const ColorPicker = ({ transactionId, currentColor, disabled }) => {
        const colors = ['green', 'blue', 'yellow', 'pink'];
        const colorTooltips = { green: '緑', blue: '青', yellow: '黄', pink: 'ピンク' };

        const getBorderClass = (color) => {
            switch (color) {
                case 'green': return 'border-green-500';
                case 'blue': return 'border-blue-500';
                case 'yellow': return 'border-yellow-500';
                case 'pink': return 'border-pink-500';
                default: return 'border-transparent';
            }
        };

        const getBgClass = (color) => {
            switch (color) {
                case 'green': return 'bg-green-400';
                case 'blue': return 'bg-blue-400';
                case 'yellow': return 'bg-yellow-400';
                case 'pink': return 'bg-pink-400';
                default: return 'bg-gray-400';
            }
        };

        return (
            <div className="flex items-center justify-center space-x-1">
                {colors.map((color) => (
                    <button
                        key={color}
                        type="button"
                        disabled={disabled}
                        onClick={() => !disabled && onColorChange(transactionId, color === currentColor ? null : color)}
                        className={`w-4 h-4 rounded-full border-2 transition-transform ${disabled ? 'opacity-40 cursor-not-allowed' : 'transform hover:scale-125'} ${
                            currentColor === color ? getBorderClass(color) : 'border-transparent'
                        } ${getBgClass(color)}`}
                        title={disabled ? 'このビューでは色付けを利用できません' : `${colorTooltips[color]}でマーク`}
                    />
                ))}
            </div>
        );
    };

    const totalWithdrawal = transactions.reduce((sum, t) => sum + (t.withdrawal || 0), 0);
    const totalDeposit = transactions.reduce((sum, t) => sum + (t.deposit || 0), 0);
    const balance = totalDeposit - totalWithdrawal;
    const showReorderColumn = showAccountInfo && typeof onReorder === 'function';
    const totalColumns = 7 + (showAccountInfo ? 2 : 0) + (showReorderColumn ? 1 : 0);
    const isSortableField = (field) => Array.isArray(sortableFields) && sortableFields.includes(field);

    const renderHeaderCell = (field, label, align = 'text-left') => {
        const sortable = typeof onSortField === 'function' && isSortableField(field);
        if (!sortable) {
            return (
                <th className={`px-3 py-3 ${align} text-xs font-medium text-gray-500 uppercase tracking-wider`}>
                    {label}
                </th>
            );
        }
        const isActive = sorting?.field === field;
        const icon = !isActive ? (
            <ArrowUpDown size={14} className="text-gray-400" />
        ) : sorting.direction === 'asc' ? (
            <ChevronUp size={14} className="text-blue-600" />
        ) : (
            <ChevronDown size={14} className="text-blue-600" />
        );
        return (
            <th className={`px-3 py-3 ${align} text-xs font-medium text-gray-500 uppercase tracking-wider`}>
                <button
                    type="button"
                    onClick={() => onSortField(field)}
                    className="inline-flex items-center gap-1 text-gray-700 hover:text-blue-600"
                >
                    <span>{label}</span>
                    {icon}
                </button>
            </th>
        );
    };

    return (
        <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                    <tr>
                        {renderHeaderCell('date', '日付')}
                        {showAccountInfo && renderHeaderCell('account', '取引口座')}
                        <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">種別</th>
                        {renderHeaderCell('withdrawal', '出金額', 'text-right')}
                        {renderHeaderCell('deposit', '入金額', 'text-right')}
                        {renderHeaderCell('memo', '備考/カテゴリ')}
                        {showAccountInfo && (
                            <th className="px-3 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">着色</th>
                        )}
                        {showReorderColumn && (
                            <th className="px-1 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">順序</th>
                        )}
                        <th className="px-3 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider" colSpan="2">操作</th>
                    </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                    {transactions.length === 0 ? (
                        <tr>
                            <td colSpan={totalColumns} className="px-6 py-4 whitespace-nowrap text-sm text-gray-500 text-center">
                                取引履歴はありません。
                            </td>
                        </tr>
                    ) : (
                        transactions.map((t, index) => {
                            const account = accountMap[t.accountId];
                            const holderName = resolveHolderName(account);
                            const getRowClass = (color) => {
                                switch (color) {
                                    case 'green':
                                        return 'bg-green-50 hover:bg-green-100';
                                    case 'blue':
                                        return 'bg-blue-50 hover:bg-blue-100';
                                    case 'yellow':
                                        return 'bg-yellow-50 hover:bg-yellow-100';
                                    case 'pink':
                                        return 'bg-pink-50 hover:bg-pink-100';
                                    default:
                                        return 'hover:bg-gray-50';
                                }
                            };
                            const rowClass = `${getRowClass(t.rowColor)} ${highlightAccountIds?.includes(t.accountId) ? 'ring-1 ring-blue-200' : ''}`;
                            return (
                                <tr key={t.id} className={`${rowClass} transition duration-150`}>
                                    <td className="px-3 py-3 whitespace-nowrap text-sm font-medium text-gray-900">{t.date}</td>
                                    {showAccountInfo && (
                                        <td className="px-3 py-3 whitespace-nowrap text-xs text-gray-600">
                                            <p className="text-[0.7rem] font-semibold text-gray-500 uppercase tracking-wide">名義人</p>
                                            <p className="text-sm text-gray-900">{holderName || '---'}</p>
                                            <p className="text-[0.7rem] font-semibold text-gray-500 uppercase tracking-wide mt-2">口座名</p>
                                            <p className="font-semibold text-gray-800">{account?.name || '不明な口座'}</p>
                                            <p className="font-mono text-gray-400 mt-1">No. {account?.number || '---'}</p>
                                        </td>
                                    )}
                                    <td className="px-3 py-3 whitespace-nowrap text-sm text-gray-700">{t.type || '---'}</td>
                                    <td className="px-3 py-3 whitespace-nowrap text-sm text-right text-red-600 font-mono">
                                        {(t.withdrawal || 0) > 0 ? formatCurrency(t.withdrawal) : '-'}
                                    </td>
                                    <td className="px-3 py-3 whitespace-nowrap text-sm text-right text-green-600 font-mono">
                                        {(t.deposit || 0) > 0 ? formatCurrency(t.deposit) : '-'}
                                    </td>
                                    <td className="px-3 py-3 whitespace-nowrap text-sm text-gray-700">
                                        {t.memo || '---'}
                                    </td>
                                    {showAccountInfo && (
                                        <td className="px-3 py-3 whitespace-nowrap text-center text-sm text-gray-500">
                                            <ColorPicker transactionId={t.id} currentColor={t.rowColor} disabled={typeof onColorChange !== 'function'} />
                                        </td>
                                    )}
                                    {showReorderColumn && (
                                        <td className="px-1 py-3 whitespace-nowrap text-center">
                                            <div className="flex flex-col items-center">
                                                {index > 0 && (
                                                    <button
                                                        onClick={() => onReorder(t, transactions[index - 1])}
                                                        title="上へ移動"
                                                        className="text-gray-500 hover:text-blue-600 p-0.5 rounded-full hover:bg-blue-50 transition"
                                                    >
                                                        <ChevronUp size={16} />
                                                    </button>
                                                )}
                                                {index < transactions.length - 1 && (
                                                    <button
                                                        onClick={() => onReorder(t, transactions[index + 1])}
                                                        title="下へ移動"
                                                        className="text-gray-500 hover:text-blue-600 p-0.5 rounded-full hover:bg-blue-50 transition"
                                                    >
                                                        <ChevronDown size={16} />
                                                    </button>
                                                )}
                                            </div>
                                        </td>
                                    )}

                                    {/* 編集ボタン */}
                                    <td className="px-3 py-3 whitespace-nowrap text-center text-sm font-medium">
                                        <button
                                            onClick={() => onEdit({ ...t, accountName: account?.name })}
                                            className="text-indigo-500 hover:text-indigo-700 p-1 rounded-full hover:bg-indigo-100 transition"
                                            title="取引を編集"
                                        >
                                            <Edit size={16} />
                                        </button>
                                    </td>

                                    {/* 削除ボタン */}
                                    <td className="px-3 py-3 whitespace-nowrap text-center text-sm font-medium">
                                        <button
                                            onClick={() => onDelete(t.id)}
                                            className="text-red-500 hover:text-red-700 p-1 rounded-full hover:bg-red-100 transition"
                                            title="取引を削除"
                                        >
                                            <Trash2 size={16} />
                                        </button>
                                    </td>
                                </tr>
                            );
                        })
                    )}
                </tbody>
                <tfoot className="bg-gray-100 font-bold">
                    <tr>
                        <td colSpan={showAccountInfo ? 3 : 2} className="px-3 py-3 text-left text-base text-gray-800">合計</td>
                        <td className="px-3 py-3 text-right text-red-600 text-base font-mono">{formatCurrency(totalWithdrawal)}</td>
                        <td className="px-3 py-3 text-right text-green-600 text-base font-mono">{formatCurrency(totalDeposit)}</td>
                        <td className="px-3 py-3 text-right text-base text-gray-800" colSpan={Math.max(1, totalColumns - (showAccountInfo ? 5 : 4))}>
                            差引残高: <span className={balance >= 0 ? 'text-green-700' : 'text-red-700'}>{formatCurrency(balance)}</span>
                        </td>
                    </tr>
                </tfoot>
            </table>
        </div>
    );
};


// --- Integrated Tab Content Component ---

const IntegratedTabContent = ({
    allTransactions,
    allAccounts,
    setEditingTransaction,
    onReorderTransactions,
    onDeleteTransaction,
    onUpdateTransactionColor,
    sorting,
    onSortChange,
    filters,
    onFilterChange,
    onFilterReset,
    comparisonSelection,
    onToggleComparisonAccount,
    onClearComparisonSelection,
    onApplyComparisonFilter,
}) => {
    const [message, setMessage] = useState(''); // メッセージ表示用ステートを追加
    const [analysisStatus, setAnalysisStatus] = useState('idle');
    const [analysisResult, setAnalysisResult] = useState(null);
    const accountMap = useMemo(() => {
        return (allAccounts || []).reduce((map, account) => {
            map[account.id] = account;
            return map;
        }, {});
    }, [allAccounts]);
    const comparisonSet = useMemo(() => new Set(comparisonSelection || []), [comparisonSelection]);
    const accountFilterOptions = useMemo(() => {
        const list = [{ value: 'all', label: 'すべての口座' }];
        if ((comparisonSelection || []).length) {
            list.push({ value: 'selected', label: '比較中の口座' });
        } else {
            list.push({ value: 'selected', label: '比較中の口座 (未選択)', disabled: true });
        }
        list.push(
            ...(allAccounts || []).map((account) => ({
                value: account.id,
                label: account.name || account.number || account.id,
            })),
        );
        return list;
    }, [allAccounts, comparisonSelection]);
    const colorOptions = [
        { value: 'all', label: 'すべて' },
        { value: 'green', label: '緑' },
        { value: 'blue', label: '青' },
        { value: 'yellow', label: '黄' },
        { value: 'pink', label: 'ピンク' },
    ];

    // 統合ロジック: userOrderがあればそれを優先し、なければ日付（昇順）でソート
    const integratedTransactions = useMemo(() => {
        return [...allTransactions].sort((a, b) => {
            // userOrderが存在しない場合、フォールバックとして日付のタイムスタンプを使用
            const orderA = a.userOrder ?? new Date(a.date).getTime();
            const orderB = b.userOrder ?? new Date(b.date).getTime();

            // 順序が異なる場合は、順序でソート
            if (orderA !== orderB) {
                return orderA - orderB;
            }
            
            // 順序が同じ（または両方未定義で日付も同じ）場合は、IDで安定ソート
            return a.id.localeCompare(b.id);
        });
    }, [allTransactions]);

    const filteredTransactions = useMemo(() => {
        if (!filters) {
            return integratedTransactions;
        }
        const keyword = (filters.keyword || '').trim().toLowerCase();
        const minAmount = filters.minAmount ? parseInt(filters.minAmount, 10) : null;
        const maxAmount = filters.maxAmount ? parseInt(filters.maxAmount, 10) : null;
        return integratedTransactions.filter((transaction) => {
            const account = accountMap[transaction.accountId];
            if (filters.accountId === 'selected') {
                if (!comparisonSet.size || !comparisonSet.has(transaction.accountId)) {
                    return false;
                }
            } else if (filters.accountId && filters.accountId !== 'all' && transaction.accountId !== filters.accountId) {
                return false;
            }
            const isWithdrawal = (transaction.withdrawal || 0) > 0;
            const isDeposit = (transaction.deposit || 0) > 0;
            if (filters.direction === 'withdrawal' && !isWithdrawal) {
                return false;
            }
            if (filters.direction === 'deposit' && !isDeposit) {
                return false;
            }
            const colorValue = transaction.rowColor || transaction.row_color || null;
            if (filters.rowColor && filters.rowColor !== 'all' && colorValue !== filters.rowColor) {
                return false;
            }
            if (keyword) {
                const targetText = [
                    transaction.memo || '',
                    transaction.type || '',
                    account?.name || '',
                    resolveHolderName(account) || '',
                ]
                    .join(' ')
                    .toLowerCase();
                if (!targetText.includes(keyword)) {
                    return false;
                }
            }
            const amountValue = Math.max(transaction.deposit || 0, transaction.withdrawal || 0);
            if (minAmount !== null && Number.isFinite(minAmount) && amountValue < minAmount) {
                return false;
            }
            if (maxAmount !== null && Number.isFinite(maxAmount) && amountValue > maxAmount) {
                return false;
            }
            return true;
        });
    }, [integratedTransactions, filters, accountMap]);

    const displayTransactions = useMemo(
        () => sortTransactionsByConfig(filteredTransactions, sorting, accountMap),
        [filteredTransactions, sorting, accountMap],
    );
    const accountSummaries = useMemo(() => {
        const totals = {};
        (allTransactions || []).forEach((txn) => {
            const accountId = txn.accountId;
            if (!accountId) {
                return;
            }
            if (!totals[accountId]) {
                totals[accountId] = {
                    deposit: 0,
                    withdrawal: 0,
                    count: 0,
                    latestTimestamp: null,
                    latestDate: '-',
                };
            }
            const record = totals[accountId];
            record.deposit += txn.deposit || 0;
            record.withdrawal += txn.withdrawal || 0;
            record.count += 1;
            if (txn.date) {
                const timestamp = new Date(txn.date).getTime();
                if (!record.latestTimestamp || timestamp > record.latestTimestamp) {
                    record.latestTimestamp = timestamp;
                    record.latestDate = txn.date;
                }
            }
        });
        return (allAccounts || []).map((account) => {
            const summary = totals[account.id] || {};
            return {
                id: account.id,
                name: account.name,
                number: account.number,
                holderName: resolveHolderName(account),
                deposit: summary.deposit || 0,
                withdrawal: summary.withdrawal || 0,
                count: summary.count || 0,
                latestDate: summary.latestDate || '-',
            };
        });
    }, [allAccounts, allTransactions]);
    const comparisonCards = useMemo(() => {
        if ((comparisonSelection || []).length) {
            return accountSummaries.filter((summary) => comparisonSelection.includes(summary.id));
        }
        return accountSummaries.slice(0, Math.min(3, accountSummaries.length));
    }, [accountSummaries, comparisonSelection]);
    const handleRunAnalysis = useCallback(() => {
        setAnalysisStatus('running');
        setTimeout(() => {
            const result = runComplianceAnalysis(allAccounts, allTransactions);
            setAnalysisResult(result);
            setAnalysisStatus('done');
        }, 40);
    }, [allAccounts, allTransactions]);
    const allowManualReorder = !sorting || sorting.field === 'custom';
    const sortOptions = [
        { value: 'custom', label: '手動順序（既定）' },
        { value: 'date', label: '日付' },
        { value: 'withdrawal', label: '出金額' },
        { value: 'deposit', label: '入金額' },
        { value: 'memo', label: '摘要' },
        { value: 'account', label: '口座名' },
    ];

    const handleReorderTransaction = async (movedItem, adjacentItem) => {
        const sourceList = displayTransactions;
        try {
            const movedIndex = sourceList.findIndex(t => t.id === movedItem.id);
            const adjacentIndex = sourceList.findIndex(t => t.id === adjacentItem.id);

            const isMovingDown = movedIndex < adjacentIndex;

            let prevItem, nextItem;

            if (isMovingDown) {
                prevItem = adjacentItem;
                nextItem = sourceList[adjacentIndex + 1];
            } else {
                prevItem = sourceList[adjacentIndex - 1];
                nextItem = adjacentItem;
            }

            const prevOrder = prevItem ? (prevItem.userOrder ?? new Date(prevItem.date).getTime()) : null;
            const nextOrder = nextItem ? (nextItem.userOrder ?? new Date(nextItem.date).getTime()) : null;
            
            let newOrder;

            if (prevOrder !== null && nextOrder !== null) {
                newOrder = (prevOrder + nextOrder) / 2;
            } else if (prevOrder !== null) {
                newOrder = prevOrder + 1000;
            } else if (nextOrder !== null) {
                newOrder = nextOrder - 1000;
            } else {
                return; // Should not happen in a list with >1 item
            }

            // --- Collision Detection & Re-balancing ---
            const isCollision = (prevOrder !== null && newOrder <= prevOrder) || (nextOrder !== null && newOrder >= nextOrder);

            if (isCollision) {
                setMessage('順序を再整理しています...');
                const reorderedTransactions = [...sourceList];
                const itemToMove = reorderedTransactions.splice(movedIndex, 1)[0];
                const newAdjacentIndex = reorderedTransactions.findIndex(t => t.id === adjacentItem.id);
                if (isMovingDown) {
                    reorderedTransactions.splice(newAdjacentIndex + 1, 0, itemToMove);
                } else {
                    reorderedTransactions.splice(newAdjacentIndex, 0, itemToMove);
                }
                const payload = reorderedTransactions.map((transaction, index) => ({
                    id: transaction.id,
                    userOrder: (index + 1) * 1000,
                }));
                await onReorderTransactions(payload);
                setMessage('順序を更新しました。');
            } else {
                await onReorderTransactions([{ id: movedItem.id, userOrder: newOrder }]);
                setMessage('順序を更新しました。');
            }
            setTimeout(() => setMessage(''), 1500);
        } catch (e) {
            console.error('Error reordering transaction:', e);
            setMessage(`順序の変更中にエラーが発生しました: ${e.message}`);
            setTimeout(() => setMessage(''), 3000);
        }
    };


    const handleColorChange = async (transactionId, color) => {
        try {
            await onUpdateTransactionColor(transactionId, color);
        } catch (e) {
            console.error('Error updating transaction color:', e);
            setMessage(`色の更新エラー: ${e.message}`);
            setTimeout(() => setMessage(''), 3000);
        }
    };

    const handleDeleteTransaction = async (transactionId) => {
        // カスタムモーダルを使う代わりに、一時的にwindow.confirmを使用
        if (!window.confirm('統合タブから削除すると、元の口座のデータも完全に削除されます。よろしいですか？')) return;

        try {
            await onDeleteTransaction(transactionId);
            setMessage('取引を削除しました。');
        } catch (e) {
            console.error('Error deleting transaction:', e);
            setMessage(`取引削除エラー: ${e.message}`);
        }
        setTimeout(() => setMessage(''), 3000);
    };

    const handleSortFromHeader = (field) => {
        if (!onSortChange) return;
        onSortChange((prev) => {
            if (field === 'custom') {
                return { field: 'custom', direction: 'asc' };
            }
            if (!prev || prev.field !== field) {
                return { field, direction: field === 'withdrawal' ? 'desc' : 'asc' };
            }
            return { field, direction: prev.direction === 'asc' ? 'desc' : 'asc' };
        });
    };

    const handleSortFieldSelect = (value) => {
        if (!onSortChange) return;
        onSortChange((prev) => {
            if (value === 'custom') {
                return { field: 'custom', direction: 'asc' };
            }
            if (!prev || prev.field !== value) {
                return { field: value, direction: value === 'withdrawal' ? 'desc' : 'asc' };
            }
            return prev;
        });
    };

    const handleSortDirectionToggle = () => {
        if (!onSortChange) return;
        onSortChange((prev) => {
            if (!prev || prev.field === 'custom') {
                return { field: 'date', direction: 'asc' };
            }
            return { field: prev.field, direction: prev.direction === 'asc' ? 'desc' : 'asc' };
        });
    };

    return (
        <div className="p-4 space-y-6" id="integrated-transactions-area">
            <div className="flex justify-between items-center border-b pb-2">
                <h2 className="text-2xl font-bold text-gray-800 flex items-center space-x-2">
                    <ArrowDownUp size={24} className="text-purple-600" />
                    <span>統合取引一覧 (手動並べ替え対応)</span>
                </h2>
                <ExportPDFButton 
                    elementId="integrated-transactions-area"
                    fileName={`統合取引一覧_${new Date().toISOString().substring(0, 10)}`}
                />
            </div>
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
                <div className="flex flex-wrap items-center gap-3">
                    <p className="text-sm text-slate-600 flex items-center gap-2 font-medium">
                        <ArrowUpDown size={16} /> 並び替えモード
                    </p>
                    <select
                        value={sorting?.field || 'custom'}
                        onChange={(e) => handleSortFieldSelect(e.target.value)}
                        className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white text-sm"
                    >
                        {sortOptions.map((option) => (
                            <option key={option.value} value={option.value}>{option.label}</option>
                        ))}
                    </select>
                    <button
                        type="button"
                        onClick={handleSortDirectionToggle}
                        disabled={!sorting || sorting.field === 'custom'}
                        className={`inline-flex items-center gap-1 px-3 py-1.5 rounded-lg border text-sm ${sorting && sorting.field !== 'custom' ? 'border-slate-400 text-slate-700 hover:bg-white/60' : 'border-slate-200 text-slate-400 cursor-not-allowed bg-white'}`}
                    >
                        {sorting?.direction === 'desc' ? '降順' : '昇順'}
                    </button>
                </div>
                <p className="text-xs text-slate-500">
                    ソート対象を切り替えると即座に表示が更新されます。手動順序モードのときのみ列内の矢印から微調整できます。
                </p>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                        <p className="text-base font-semibold text-slate-800">AI分析 (ヒューリスティック)</p>
                        <p className="text-xs text-slate-500">保険契約の有無や贈与税リスクなど、典型的な論点を自動チェックします。</p>
                    </div>
                    <div className="flex items-center gap-2">
                        {analysisStatus === 'running' && <Loader2 size={18} className="animate-spin text-slate-500" />}
                        <button
                            type="button"
                            onClick={handleRunAnalysis}
                            className="rounded-lg bg-[#2563eb] px-4 py-2 text-sm font-semibold text-white hover:bg-[#1d4ed8]"
                            disabled={analysisStatus === 'running'}
                        >
                            AI分析を実行
                        </button>
                    </div>
                </div>
                {analysisResult && (
                    <div className="space-y-3">
                        <p className="text-xs text-slate-400">更新: {new Date(analysisResult.generatedAt).toLocaleString()}</p>
                        {analysisResult.checks.map((check, index) => (
                            <div key={`${check.category}-${index}`} className="rounded-xl border border-slate-200 p-3 bg-slate-50">
                                <div className="flex items-center justify-between">
                                    <h4 className="text-sm font-semibold text-slate-800">{check.category}</h4>
                                    <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${check.severity === 'warn' ? 'bg-amber-100 text-amber-800' : 'bg-blue-100 text-blue-700'}`}>
                                        {check.severity === 'warn' ? '確認' : '参考'}
                                    </span>
                                </div>
                                <p className="text-sm text-slate-600 mt-1">{check.message}</p>
                                {check.items && check.items.length > 0 && (
                                    <ul className="mt-2 text-xs text-slate-500 list-disc list-inside space-y-1">
                                        {check.items.map((item, itemIndex) => (
                                            <li key={`${check.category}-${index}-${itemIndex}`}>{item}</li>
                                        ))}
                                    </ul>
                                )}
                            </div>
                        ))}
                    </div>
                )}
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-4">
                <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                        <h3 className="text-base font-semibold text-slate-800">口座比較パネル</h3>
                        <p className="text-xs text-slate-500">比較したい口座をタップすると下のカードに並べて表示されます。</p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                        <button
                            type="button"
                            onClick={onApplyComparisonFilter}
                            disabled={!comparisonSelection?.length}
                            className={`rounded-lg px-3 py-1.5 text-sm font-semibold ${comparisonSelection?.length ? 'bg-blue-600 text-white hover:bg-blue-700' : 'bg-slate-200 text-slate-500 cursor-not-allowed'}`}
                        >
                            比較中の口座で表示
                        </button>
                        <button
                            type="button"
                            onClick={onClearComparisonSelection}
                            disabled={!comparisonSelection?.length}
                            className={`rounded-lg px-3 py-1.5 text-sm ${comparisonSelection?.length ? 'border border-slate-300 text-slate-600 hover:bg-slate-50' : 'border border-slate-200 text-slate-400 cursor-not-allowed'}`}
                        >
                            比較をクリア
                        </button>
                    </div>
                </div>
                <div className="flex flex-wrap gap-2">
                    {accountSummaries.map((summary) => {
                        const isSelected = comparisonSelection?.includes(summary.id);
                        return (
                            <button
                                key={summary.id}
                                type="button"
                                onClick={() => onToggleComparisonAccount(summary.id)}
                                className={`rounded-full border px-3 py-1 text-xs font-semibold transition ${isSelected ? 'bg-blue-600 text-white border-blue-600' : 'border-slate-300 text-slate-600 hover:bg-slate-50'}`}
                            >
                                {summary.name || summary.number || '口座'}
                            </button>
                        );
                    })}
                    {accountSummaries.length === 0 && <p className="text-xs text-slate-500">口座がありません。</p>}
                </div>
                <div className="grid gap-3 md:grid-cols-2">
                    {comparisonCards.length === 0 && (
                        <p className="text-sm text-slate-500">比較したい口座を上のボタンから選択してください。</p>
                    )}
                    {comparisonCards.map((summary) => (
                        <div key={summary.id} className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                            <div className="flex items-center justify-between">
                                <div>
                                    <p className="text-xs uppercase tracking-widest text-slate-500">{summary.number || 'No. ---'}</p>
                                    <p className="text-lg font-semibold text-slate-900">{summary.name || '口座'}</p>
                                    <p className="text-xs text-slate-500">{summary.holderName || '名義未登録'}</p>
                                </div>
                                <div className="text-right text-xs text-slate-500">
                                    <p>取引 {summary.count || 0} 件</p>
                                    <p>最終更新 {summary.latestDate}</p>
                                </div>
                            </div>
                            <div className="mt-4 grid grid-cols-3 gap-3 text-center">
                                <div>
                                    <p className="text-[0.65rem] uppercase text-slate-500">出金</p>
                                    <p className="text-sm font-semibold text-red-600">{formatCurrency(summary.withdrawal || 0)}</p>
                                </div>
                                <div>
                                    <p className="text-[0.65rem] uppercase text-slate-500">入金</p>
                                    <p className="text-sm font-semibold text-green-600">{formatCurrency(summary.deposit || 0)}</p>
                                </div>
                                <div>
                                    <p className="text-[0.65rem] uppercase text-slate-500">差引</p>
                                    <p className={`text-sm font-semibold ${((summary.deposit || 0) - (summary.withdrawal || 0)) >= 0 ? 'text-green-700' : 'text-red-700'}`}>
                                        {formatCurrency((summary.deposit || 0) - (summary.withdrawal || 0))}
                                    </p>
                                </div>
                            </div>
                        </div>
                    ))}
                </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-4">
                <div className="grid gap-3 md:grid-cols-3">
                    <div>
                        <label className="text-xs text-slate-500">口座で絞り込む</label>
                        <select
                            value={filters?.accountId || 'all'}
                            onChange={(e) => onFilterChange?.({ accountId: e.target.value })}
                            className="mt-1 w-full rounded-lg border border-slate-300 p-2 text-sm"
                        >
                            {accountFilterOptions.map((option) => (
                                <option key={option.value} value={option.value} disabled={option.disabled}>
                                    {option.label}
                                </option>
                            ))}
                        </select>
                    </div>
                    <div>
                        <label className="text-xs text-slate-500">入出金区分</label>
                        <select
                            value={filters?.direction || 'all'}
                            onChange={(e) => onFilterChange?.({ direction: e.target.value })}
                            className="mt-1 w-full rounded-lg border border-slate-300 p-2 text-sm"
                        >
                            <option value="all">すべて</option>
                            <option value="deposit">入金のみ</option>
                            <option value="withdrawal">出金のみ</option>
                        </select>
                    </div>
                    <div>
                        <label className="text-xs text-slate-500">着色フィルター</label>
                        <select
                            value={filters?.rowColor || 'all'}
                            onChange={(e) => onFilterChange?.({ rowColor: e.target.value })}
                            className="mt-1 w-full rounded-lg border border-slate-300 p-2 text-sm"
                        >
                            {colorOptions.map((option) => (
                                <option key={option.value} value={option.value}>{option.label}</option>
                            ))}
                        </select>
                    </div>
                    <div>
                        <label className="text-xs text-slate-500">最小金額</label>
                        <input
                            type="number"
                            inputMode="numeric"
                            value={filters?.minAmount || ''}
                            onChange={(e) => onFilterChange?.({ minAmount: e.target.value })}
                            className="mt-1 w-full rounded-lg border border-slate-300 p-2 text-sm"
                            placeholder="0"
                        />
                    </div>
                    <div>
                        <label className="text-xs text-slate-500">最大金額</label>
                        <input
                            type="number"
                            inputMode="numeric"
                            value={filters?.maxAmount || ''}
                            onChange={(e) => onFilterChange?.({ maxAmount: e.target.value })}
                            className="mt-1 w-full rounded-lg border border-slate-300 p-2 text-sm"
                            placeholder="0"
                        />
                    </div>
                    <div>
                        <label className="text-xs text-slate-500">キーワード</label>
                        <input
                            type="text"
                            value={filters?.keyword || ''}
                            onChange={(e) => onFilterChange?.({ keyword: e.target.value })}
                            className="mt-1 w-full rounded-lg border border-slate-300 p-2 text-sm"
                            placeholder="メモ・摘要で検索"
                        />
                    </div>
                </div>
                <div className="flex justify-end gap-3">
                    <button
                        type="button"
                        onClick={onFilterReset}
                        className="rounded-lg border border-slate-300 px-4 py-2 text-sm text-slate-600 hover:bg-slate-50"
                    >
                        条件をクリア
                    </button>
                </div>
            </div>

            {/* メッセージ表示 */}
            {message && <p className={`p-3 rounded-lg my-3 text-sm flex items-center justify-center ${message.includes('エラー') ? 'bg-red-100 text-red-700' : 'bg-blue-100 text-blue-700'}`}><Loader2 size={16} className="animate-spin mr-2" /> {message}</p>}

            <div className="bg-white rounded-xl shadow-lg overflow-hidden">
                <TransactionTable
                    transactions={displayTransactions}
                    accounts={allAccounts}
                    onDelete={handleDeleteTransaction}
                    onEdit={setEditingTransaction}
                    onReorder={allowManualReorder ? handleReorderTransaction : undefined}
                    onColorChange={handleColorChange}
                    showAccountInfo={true}
                    highlightAccountIds={comparisonSelection}
                    sorting={sorting}
                    onSortField={handleSortFromHeader}
                    sortableFields={['date', 'account', 'withdrawal', 'deposit', 'memo']}
                />
                {!allowManualReorder && (
                    <p className="px-4 py-2 text-xs text-slate-500 bg-slate-50 border-t border-slate-100">※ 並び替えモードが「手動順序」以外の間は順序列のボタンを無効化しています。</p>
                )}
            </div>
        </div>
    );
};

// --- Main App Component ---

const LedgerApp = () => {
    const [sessionToken, setSessionToken] = useState(null);
    const [userId, setUserId] = useState(null);
    const [cases, setCases] = useState([]);
    const [selectedCaseId, setSelectedCaseId] = useState(null);
    const [accounts, setAccounts] = useState([]);
    const [transactions, setTransactions] = useState([]);
    const [activeTab, setActiveTab] = useState('register');
    const [showAddAccountModal, setShowAddAccountModal] = useState(false);
    const [showExportModal, setShowExportModal] = useState(false);
    const [showImportModal, setShowImportModal] = useState(false);
    const [editingTransaction, setEditingTransaction] = useState(null);
    const [editingAccount, setEditingAccount] = useState(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState(null);
    const [jobPreview, setJobPreview] = useState(null);
    const [jobImportMappings, setJobImportMappings] = useState({});
    const [jobImportStatus, setJobImportStatus] = useState('idle');
    const [jobImportError, setJobImportError] = useState('');
    const [newCaseName, setNewCaseName] = useState('');
    const [pendingImports, setPendingImports] = useState([]);
    const [showPendingImportModal, setShowPendingImportModal] = useState(false);
    const [pendingImportStatus, setPendingImportStatus] = useState('idle');
    const [pendingImportError, setPendingImportError] = useState('');
    const [showGuide, setShowGuide] = useState(false);
    const [transactionSort, setTransactionSort] = useState({ field: 'custom', direction: 'asc' });
    const [transactionFilter, setTransactionFilter] = useState(() => ({ ...INITIAL_TRANSACTION_FILTER }));
    const [comparisonSelection, setComparisonSelection] = useState([]);
    const accountFilterMode = transactionFilter.accountId;

    useEffect(() => {
        if (comparisonSelection.length === 0 && accountFilterMode === 'selected') {
            setTransactionFilter((prev) => ({
                ...prev,
                accountId: 'all',
            }));
        }
    }, [comparisonSelection, accountFilterMode]);
    const initialJobId = useMemo(() => {
        const params = new URLSearchParams(window.location.search);
        return params.get('job_id');
    }, []);
    const initialCaseIdRef = useRef(new URLSearchParams(window.location.search).get('case_id'));

    const callLedgerApi = useCallback(async (path, options = {}) => {
        if (!sessionToken) {
            throw new Error('セッションが初期化されていません。');
        }
        const headers = {
            'Content-Type': 'application/json',
            'X-Ledger-Token': sessionToken,
            'X-Ledger-App': appId,
            ...(options.headers || {}),
        };
        const init = {
            method: options.method || 'GET',
            ...options,
            headers,
        };
        if (options.body !== undefined) {
            init.body = typeof options.body === 'string' ? options.body : JSON.stringify(options.body);
        }
        const url = path.startsWith('http') ? path : `${ledgerApiBase}${path}`;
        const response = await fetch(url, init);
        if (!response.ok) {
            let message = 'Ledger API error';
            try {
                const data = await response.json();
                message = data.detail || data.message || message;
            } catch {
                const text = await response.text();
                if (text) message = text;
            }
            throw new Error(message);
        }
        if (response.status === 204) {
            return null;
        }
        return response.json();
    }, [sessionToken, ledgerApiBase]);

    const refreshState = useCallback(
        async (caseIdOverride = null, showSpinner = true) => {
            if (!sessionToken) return;
            const targetCaseId = caseIdOverride || selectedCaseId;
            if (!targetCaseId) return;
            if (showSpinner) setLoading(true);
            try {
                const data = await callLedgerApi(`/state?case_id=${encodeURIComponent(targetCaseId)}`);
                setAccounts(data.accounts || []);
                setTransactions(data.transactions || []);
                if (Array.isArray(data.cases)) {
                    setCases(data.cases);
                }
                if (data.case?.id) {
                    setSelectedCaseId(data.case.id);
                }
                setError(null);
            } catch (err) {
                console.error('Failed to fetch ledger state:', err);
                setError(err);
            } finally {
                if (showSpinner) setLoading(false);
            }
        },
        [sessionToken, selectedCaseId, callLedgerApi],
    );

    const fetchCases = useCallback(async () => {
        if (!sessionToken) return [];
        const data = await callLedgerApi('/cases');
        const fetchedCases = data.cases || [];
        setCases(fetchedCases);
        setSelectedCaseId((current) => current || initialCaseIdRef.current || fetchedCases[0]?.id || null);
        return fetchedCases;
    }, [sessionToken, callLedgerApi]);

    useEffect(() => {
        let cancelled = false;
        const bootstrap = async () => {
            setLoading(true);
            try {
                const storedToken = getStoredLedgerToken();
                const payload = { app_id: appId };
                if (storedToken) {
                    payload.session_token = storedToken;
                }
                const response = await fetch(`${ledgerApiBase}/session`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload),
                });
                if (!response.ok) {
                    throw new Error('Ledger API セッションの初期化に失敗しました。');
                }
                const data = await response.json();
                if (!cancelled) {
                    persistLedgerToken(data.session_token);
                    setSessionToken(data.session_token);
                    setUserId(data.user_id);
                    setError(null);
                }
            } catch (err) {
                if (!cancelled) {
                    console.error('Ledger session initialization failed:', err);
                    setError(err);
                    setLoading(false);
                }
            }
        };
        bootstrap();
        return () => {
            cancelled = true;
        };
    }, []);

    useEffect(() => {
        if (!sessionToken) return;
        (async () => {
            const list = await fetchCases();
            const targetCaseId = initialCaseIdRef.current || list[0]?.id;
            if (targetCaseId) {
                await refreshState(targetCaseId, true);
                initialCaseIdRef.current = null;
            } else {
                setLoading(false);
            }
        })();
    }, [sessionToken, fetchCases, refreshState]);

    useEffect(() => {
        setPendingImports(loadPendingImportsFromStorage());
        const handler = (event) => {
            if (event.key === PENDING_IMPORT_STORAGE_KEY) {
                setPendingImports(loadPendingImportsFromStorage());
            }
        };
        window.addEventListener('storage', handler);
        return () => window.removeEventListener('storage', handler);
    }, []);

    const refreshPendingImports = useCallback(() => {
        const entries = loadPendingImportsFromStorage();
        setPendingImports(entries);
        return entries;
    }, []);

    const removePendingImportEntry = useCallback((entryId) => {
        const entries = loadPendingImportsFromStorage().filter((entry) => entry.id !== entryId);
        savePendingImportsToStorage(entries);
        setPendingImports(entries);
        return entries;
    }, []);

    const convertAssetsToLedgerPayload = useCallback((entry) => {
        const accounts = [];
        const transactions = [];
        (entry?.assets || []).forEach((asset, index) => {
            const identifiers = asset?.identifiers || {};
            const accountId = String(asset?.record_id || identifiers.primary || `${entry.id || 'pending'}_${index + 1}`);
            const holderName = Array.isArray(asset?.owner_name)
                ? (asset.owner_name.filter(Boolean).join(' / ') || undefined)
                : (asset?.owner_name || undefined);
            const name = asset?.asset_name || entry?.name || holderName || `口座${index + 1}`;
            accounts.push({
                id: accountId,
                name,
                number: identifiers.primary || identifiers.secondary || '',
                holderName,
                order: (index + 1) * 1000,
            });
            (asset?.transactions || []).forEach((txn, txnIndex) => {
                transactions.push({
                    id: `${accountId}-${txnIndex + 1}`,
                    accountId,
                    date: txn?.transaction_date,
                    withdrawal: parseInt(txn?.withdrawal_amount || 0, 10),
                    deposit: parseInt(txn?.deposit_amount || 0, 10),
                    memo: txn?.correction_note || txn?.memo || txn?.description || '',
                    type: txn?.description || '',
                });
            });
        });
        return { accounts, transactions };
    }, []);

    const fetchJobPreview = useCallback(
        async (jobId) => {
            if (!jobId) return;
            try {
                setJobImportStatus('loading');
                const data = await callLedgerApi(`/jobs/${jobId}/preview`);
                const normalizedAccounts = (data.accounts || []).map((account) => ({
                    ...account,
                    assetId: account.assetId || account.asset_id,
                }));
                const defaultMappings = {};
                normalizedAccounts.forEach((account) => {
                    defaultMappings[account.assetId] = {
                        mode: 'new',
                        accountName: account.accountName,
                        accountNumber: account.accountNumber,
                        holderName: (account.ownerName && account.ownerName.filter(Boolean).join(' / ')) || '',
                    };
                });
                setJobPreview({ jobId: data.job_id || data.jobId, accounts: normalizedAccounts });
                setJobImportMappings(defaultMappings);
                setJobImportStatus('idle');
            } catch (err) {
                console.error('Failed to fetch job preview:', err);
                setJobPreview(null);
                setJobImportStatus('error');
                setJobImportError(err.message);
            }
        },
        [callLedgerApi],
    );

useEffect(() => {
    if (!sessionToken || !initialJobId) return;
    fetchJobPreview(initialJobId);
}, [sessionToken, initialJobId, fetchJobPreview]);

const handleAddAccountClick = useCallback(() => {
        if (pendingImports.length > 0) {
            setShowPendingImportModal(true);
        } else {
            setShowAddAccountModal(true);
        }
    }, [pendingImports]);

    const handleManualAddAccount = useCallback(() => {
        setShowPendingImportModal(false);
        setShowAddAccountModal(true);
    }, []);

    const handleDismissPendingEntry = useCallback((entryId) => {
        removePendingImportEntry(entryId);
        setPendingImportError('');
    }, [removePendingImportEntry]);

    const handleCreateAccount = useCallback(
        async ({ name, number, holderName }) => {
            if (!selectedCaseId) {
                throw new Error('案件が選択されていません。');
            }
            await callLedgerApi('/accounts', {
                method: 'POST',
                body: { name, number, caseId: selectedCaseId, holderName },
            });
            await refreshState(selectedCaseId, false);
        },
        [callLedgerApi, refreshState, selectedCaseId],
    );

    const handleReorderAccounts = useCallback(async (items) => {
        if (!items || items.length === 0 || !selectedCaseId) return;
        await callLedgerApi(`/accounts/reorder?case_id=${encodeURIComponent(selectedCaseId)}`, {
            method: 'POST',
            body: { items },
        });
        await refreshState(selectedCaseId, false);
    }, [callLedgerApi, refreshState, selectedCaseId]);

    const handleDeleteAccount = useCallback(async (accountId) => {
        await callLedgerApi(`/accounts/${accountId}`, { method: 'DELETE' });
        await refreshState(selectedCaseId, false);
    }, [callLedgerApi, refreshState, selectedCaseId]);

    const handleUpdateAccount = useCallback(async (accountId, payload) => {
        await callLedgerApi(`/accounts/${accountId}`, {
            method: 'PATCH',
            body: payload,
        });
        await refreshState(selectedCaseId, false);
    }, [callLedgerApi, refreshState, selectedCaseId]);

    const handleCreateTransaction = useCallback(async (payload) => {
        await callLedgerApi('/transactions', {
            method: 'POST',
            body: payload,
        });
        await refreshState(selectedCaseId, false);
    }, [callLedgerApi, refreshState, selectedCaseId]);

    const handleUpdateTransaction = useCallback(async (transactionId, payload) => {
        await callLedgerApi(`/transactions/${transactionId}`, {
            method: 'PATCH',
            body: payload,
        });
        await refreshState(selectedCaseId, false);
    }, [callLedgerApi, refreshState, selectedCaseId]);

    const handleDeleteTransaction = useCallback(async (transactionId) => {
        await callLedgerApi(`/transactions/${transactionId}`, { method: 'DELETE' });
        await refreshState(selectedCaseId, false);
    }, [callLedgerApi, refreshState, selectedCaseId]);

    const handleReorderTransactions = useCallback(async (items) => {
        if (!items || items.length === 0) return;
        await callLedgerApi('/transactions/reorder', {
            method: 'POST',
            body: { items },
        });
        await refreshState(selectedCaseId, false);
    }, [callLedgerApi, refreshState, selectedCaseId]);

    const handleUpdateTransactionColor = useCallback(async (transactionId, rowColor) => {
        await callLedgerApi(`/transactions/${transactionId}`, {
            method: 'PATCH',
            body: { rowColor },
        });
        await refreshState(selectedCaseId, false);
    }, [callLedgerApi, refreshState, selectedCaseId]);

    const handleImportData = useCallback(async (payload) => {
        if (!selectedCaseId) {
            throw new Error('案件が選択されていません。');
        }
        await callLedgerApi('/import', {
            method: 'POST',
            body: { ...payload, caseId: selectedCaseId },
        });
        await refreshState(selectedCaseId, false);
    }, [callLedgerApi, refreshState, selectedCaseId]);

    const handleJobMappingChange = useCallback((assetId, updates) => {
        setJobImportMappings((prev) => ({
            ...prev,
            [assetId]: {
                ...prev[assetId],
                ...updates,
            },
        }));
    }, []);

    const handleTransactionSortChange = useCallback((updater) => {
        setTransactionSort((prev) => {
            const next = typeof updater === 'function' ? updater(prev) : updater;
            if (!next || !next.field || next.field === 'custom') {
                return { field: 'custom', direction: 'asc' };
            }
            const direction = next.direction || (next.field === 'withdrawal' ? 'desc' : 'asc');
            return { field: next.field, direction };
        });
    }, []);

    const handleTransactionFilterChange = useCallback((updates) => {
        setTransactionFilter((prev) => ({
            ...prev,
            ...updates,
        }));
    }, []);

    const handleTransactionFilterReset = useCallback(() => {
        setTransactionFilter({ ...INITIAL_TRANSACTION_FILTER });
    }, []);

    const toggleComparisonAccount = useCallback((accountId) => {
        setComparisonSelection((prev) => {
            if (prev.includes(accountId)) {
                return prev.filter((id) => id !== accountId);
            }
            return [...prev, accountId];
        });
    }, []);

    const clearComparisonSelection = useCallback(() => {
        setComparisonSelection([]);
    }, []);

    const applyComparisonFilter = useCallback(() => {
        setTransactionFilter((prev) => ({
            ...prev,
            accountId: 'selected',
        }));
    }, []);

    const handleImportPendingEntry = useCallback(
        async (entry, { targetCaseId, newCaseName } = {}) => {
            if (!entry) return;
            const payload = convertAssetsToLedgerPayload(entry);
            if (!payload.accounts.length) {
                throw new Error('取り込む口座がありません。');
            }
            const caseIdToUse = newCaseName ? null : targetCaseId || selectedCaseId;
            if (!caseIdToUse && !newCaseName) {
                throw new Error('案件が選択されていません。');
            }
            setPendingImportStatus('applying');
            setPendingImportError('');
            try {
                let requestBody = {
                    caseId: caseIdToUse,
                    newCaseName: newCaseName || null,
                    accounts: payload.accounts,
                    transactions: payload.transactions,
                };
                if (caseIdToUse && !newCaseName && selectedCaseId === caseIdToUse) {
                    requestBody = buildMergedImportPayload({
                        caseId: caseIdToUse,
                        existingAccounts: accounts,
                        existingTransactions: transactions,
                        incomingAccounts: payload.accounts,
                        incomingTransactions: payload.transactions,
                    });
                }
                await callLedgerApi('/import', {
                    method: 'POST',
                    body: requestBody,
                });
                const remaining = removePendingImportEntry(entry.id);
                if (!remaining.length) {
                    setShowPendingImportModal(false);
                }
                await fetchCases();
                const nextCaseId = caseIdToUse || selectedCaseId;
                if (nextCaseId) {
                    await refreshState(nextCaseId, true);
                }
                setPendingImportStatus('idle');
                setPendingImportError('');
            } catch (err) {
                console.error('Pending import failed:', err);
                setPendingImportStatus('error');
                setPendingImportError(err instanceof Error ? err.message : String(err));
            }
        },
        [
            callLedgerApi,
            convertAssetsToLedgerPayload,
            removePendingImportEntry,
            fetchCases,
            refreshState,
            selectedCaseId,
            accounts,
            transactions,
        ],
    );

    const changeCase = useCallback(async (caseId) => {
        if (!caseId) return;
        setSelectedCaseId(caseId);
        await refreshState(caseId, true);
    }, [refreshState]);

    const handleCaseSelectChange = useCallback(
        async (eventOrValue) => {
            const nextCaseId = typeof eventOrValue === 'string' ? eventOrValue : eventOrValue.target.value;
            await changeCase(nextCaseId);
        },
        [changeCase],
    );

    const handleCreateCaseClick = useCallback(async () => {
        const name = window.prompt('新しい案件名を入力してください', '案件');
        if (!name) {
            return;
        }
        try {
            const created = await callLedgerApi('/cases', {
                method: 'POST',
                body: { name },
            });
            await fetchCases();
            await changeCase(created.id);
        } catch (err) {
            console.error('Failed to create case:', err);
            setError(err);
        }
    }, [callLedgerApi, fetchCases, changeCase, setError]);

    const handleApplyJobImport = useCallback(async () => {
        if (!jobPreview || !jobPreview.accounts?.length) return;
        if (!selectedCaseId && !newCaseName) {
            setJobImportError('案件を選択するか、新しい案件名を入力してください。');
            return;
        }
        try {
            setJobImportStatus('applying');
            setJobImportError('');
            const payload = {
                caseId: newCaseName ? null : selectedCaseId,
                newCaseName: newCaseName || null,
                mappings: jobPreview.accounts.map((account) => {
                    const config = jobImportMappings[account.assetId] || { mode: 'new' };
                    const isGroupMode = config.mode === 'group';
                    if (config.mode === 'merge' && !config.targetAccountId) {
                        throw new Error(`${account.accountName || '口座'} のマージ先を選択してください。`);
                    }
                    if (isGroupMode && !config.mergeGroupKey) {
                        throw new Error(`${account.accountName || '口座'} の統合キーを入力してください。`);
                    }
                    const defaultHolderName = (account.ownerName && account.ownerName.filter(Boolean).join(' / ')) || undefined;
                    const holderName = config.holderName || defaultHolderName;
                    return {
                        assetId: account.assetId,
                        mode: isGroupMode ? 'new' : (config.mode || 'new'),
                        targetAccountId: isGroupMode ? null : (config.targetAccountId || null),
                        accountName: config.accountName || account.accountName,
                        accountNumber: config.accountNumber || account.accountNumber,
                        holderName,
                        groupKey: isGroupMode ? config.mergeGroupKey : null,
                        groupName: isGroupMode ? (config.mergeGroupName || config.accountName || account.accountName) : null,
                        groupNumber: isGroupMode ? (config.mergeGroupNumber || config.accountNumber || account.accountNumber) : null,
                        groupHolderName: isGroupMode ? holderName : null,
                    };
                }),
            };
            const response = await callLedgerApi(`/jobs/${jobPreview.jobId}/import`, {
                method: 'POST',
                body: payload,
            });
            const nextCaseId = response.caseId || selectedCaseId;
            await fetchCases();
            if (nextCaseId) {
                setSelectedCaseId(nextCaseId);
                await refreshState(nextCaseId, true);
            }
            setJobPreview(null);
            setJobImportMappings({});
            setNewCaseName('');
            if (initialJobId) {
                const url = new URL(window.location.href);
                url.searchParams.delete('job_id');
                window.history.replaceState({}, '', url.toString());
            }
            setJobImportStatus('success');
        } catch (err) {
            console.error('Job import failed:', err);
            setJobImportStatus('error');
            setJobImportError(err.message);
        }
    }, [jobPreview, jobImportMappings, selectedCaseId, newCaseName, initialJobId, callLedgerApi, fetchCases, refreshState]);

    // 3. アクティブなタブの内容をレンダリング
    const renderContent = () => {
        if (loading || !sessionToken) {
            return <div className="p-8"><StatusMessage loading={loading} error={error} userId={userId} /></div>;
        }
        
        if (activeTab === 'register') {
            return (
                <AccountManagementContent 
                    accounts={accounts} 
                    caseName={cases.find((item) => item.id === selectedCaseId)?.name}
                    setShowExportModal={setShowExportModal}
                    setShowImportModal={setShowImportModal}
                    onReorderAccountOrder={handleReorderAccounts}
                    onDeleteAccount={handleDeleteAccount}
                    onAddAccountClick={handleAddAccountClick}
                    onEditAccount={setEditingAccount}
                />
            );
        }

        if (activeTab === 'integrated') {
            return (
                <IntegratedTabContent
                    allTransactions={transactions}
                    allAccounts={accounts}
                    setEditingTransaction={setEditingTransaction}
                    onReorderTransactions={handleReorderTransactions}
                    onDeleteTransaction={handleDeleteTransaction}
                    onUpdateTransactionColor={handleUpdateTransactionColor}
                    sorting={transactionSort}
                    onSortChange={handleTransactionSortChange}
                    filters={transactionFilter}
                    onFilterChange={handleTransactionFilterChange}
                    onFilterReset={handleTransactionFilterReset}
                    comparisonSelection={comparisonSelection}
                    onToggleComparisonAccount={toggleComparisonAccount}
                    onClearComparisonSelection={clearComparisonSelection}
                    onApplyComparisonFilter={applyComparisonFilter}
                />
            );
        }

        const activeAccount = accounts.find(acc => acc.id === activeTab);
        if (activeAccount) {
            return <TransactionTabContent 
                        account={activeAccount} 
                        transactions={transactions} 
                        onCreateTransaction={handleCreateTransaction}
                        onDeleteTransaction={handleDeleteTransaction}
                        setEditingTransaction={setEditingTransaction}
                    />;
        }

        // 口座がまだない場合、登録を促す
        if (accounts.length === 0) {
            return (
                <div className="p-8 text-center text-gray-600">
                    <h3 className="text-xl font-semibold">はじめに、取引口座を登録しましょう</h3>
                    <p className="mt-2">上の「<span className="font-bold text-blue-600">設定・並べ替え</span>」タブ、または右上の「<span className="font-bold text-green-600">新規口座を登録</span>」ボタンから、最初の口座を追加してください。</p>
                </div>
            )
        }

        return (
            <div className="p-8 text-center text-gray-600">
                表示する口座を選択してください。
            </div>
        );
    };

    // 4. タブのレンダリング
    const TabButton = ({ tabId, label, Icon, className = '', badge = 0, number = null }) => {
        const isActive = activeTab === tabId;
        return (
            <button
                onClick={() => setActiveTab(tabId)}
                // ボタン全体のレイアウトを縦並び(flex-col)に変更
                className={`flex flex-col items-start space-y-0.5 px-4 py-3 text-sm font-semibold rounded-t-lg transition-all duration-200 focus:outline-none ${className} ${
                    isActive
                        ? 'bg-white text-blue-600 border-b-4 border-blue-600 shadow-t-lg'
                        : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                }`}
            >
                {/* 1行目: アイコン、ラベル、バッジ */}
                <div className="flex items-center space-x-2">
                    {Icon && <Icon size={18} />}
                    <span>{label}</span>
                    {badge > 0 && <span className="ml-1 text-xs bg-red-500 text-white rounded-full h-5 w-5 flex items-center justify-center">{badge}</span>}
                </div>
                {/* 2行目: 口座番号 (アイコンの幅に合わせてインデント) */}
                {number && (
                    <span className="text-xs font-normal text-gray-400 ml-[22px] -mt-1">No. {number}</span>
                )}
            </button>
        );
    };

    const totalTransactionCount = transactions.length;

    return (
        <div className="min-h-screen bg-[#f1f5f9] font-sans">
            <header className="relative overflow-hidden text-white border-b border-[#102345]" style={{ background: 'linear-gradient(135deg, #050d26 0%, #0d2765 55%, #15418f 100%)' }}>
                <div className="absolute inset-0 pointer-events-none" style={{
                    background: 'radial-gradient(circle at 12% 18%, rgba(255,255,255,0.18), transparent 55%), radial-gradient(circle at 88% 6%, rgba(96,165,250,0.32), transparent 68%)',
                }} />
                <div className="relative max-w-7xl mx-auto py-7 px-4 sm:px-6 lg:px-8 flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
                    <div>
                        <p className="text-xs uppercase tracking-[0.32em] text-[#c7d2fe]">SOROBOCR LEDGER</p>
                        <h1 className="text-3xl font-extrabold flex items-center gap-2">
                            <List size={30} className="text-[#93c5fd]" />
                            <span>入出金検討表ツール</span>
                        </h1>
                        <p className="text-sm text-[rgba(224,231,255,0.9)] mt-1 max-w-2xl">SOROBOCRで読み取ったCSV/JSONをRailway上のLedger APIに同期し、ブラウザだけで仕分け・比較・出力まで完了できます。</p>
                    </div>
                    <div className="flex flex-wrap gap-3">
                        <button
                            type="button"
                            onClick={() => setShowGuide(true)}
                            className="inline-flex items-center gap-2 rounded-full border border-white/40 px-4 py-2 text-sm font-semibold text-white hover:bg-white/10"
                        >
                            <BookOpen size={16} /> 使い方を見る
                        </button>
                        <a
                            href="./guide.html"
                            target="_blank"
                            rel="noopener"
                            className="inline-flex items-center gap-2 rounded-full border border-white/25 px-4 py-2 text-sm font-semibold text-white hover:bg-white/10"
                        >
                            詳細ガイド
                        </a>
                        <MainButton Icon={Plus} onClick={() => setShowAddAccountModal(true)} className="bg-[#22c55e] hover:bg-[#16a34a] px-4 py-2">
                            新規口座を登録
                        </MainButton>
                    </div>
                </div>
            </header>

            <div className="max-w-7xl mx-auto mt-6 px-4 sm:px-6 lg:px-8">
                <div className="flex flex-wrap items-end gap-4 bg-white border border-gray-200 rounded-xl p-4 shadow-sm mb-4">
                    <div className="flex-1 min-w-[220px]">
                        <label className="text-sm text-gray-600 mb-1 block">案件を選択</label>
                        <select
                            value={selectedCaseId || ''}
                            onChange={handleCaseSelectChange}
                            className="p-2.5 border border-slate-300 rounded-lg w-full bg-white"
                        >
                            {cases.length === 0 && <option value="">案件がありません</option>}
                            {cases.map((item) => (
                                <option key={item.id} value={item.id}>{item.name}</option>
                            ))}
                        </select>
                    </div>
                    <div>
                        <MainButton Icon={Plus} onClick={handleCreateCaseClick} className="bg-indigo-600 hover:bg-indigo-700">
                            案件を追加
                        </MainButton>
                    </div>
                </div>

                {selectedCaseId && (
                    <div className="grid gap-3 sm:grid-cols-3 mb-6">
                        <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
                            <p className="text-xs text-slate-500">案件名</p>
                            <p className="text-lg font-semibold text-slate-900">{cases.find((c) => c.id === selectedCaseId)?.name || '---'}</p>
                        </div>
                        <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
                            <p className="text-xs text-slate-500">登録口座数</p>
                            <p className="text-2xl font-bold text-slate-900">{accounts.length}</p>
                        </div>
                        <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
                            <p className="text-xs text-slate-500">取引件数</p>
                            <p className="text-2xl font-bold text-slate-900">{transactions.length}</p>
                        </div>
                    </div>
                )}

                {jobPreview && (
                    <section className="bg-yellow-50 border border-yellow-200 rounded-2xl p-5 mb-5 space-y-4">
                        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
                            <div>
                                <h3 className="text-xl font-semibold text-yellow-900">OCR結果を案件へ取り込み</h3>
                                <p className="text-sm text-yellow-800">ジョブID: {jobPreview.jobId} ／ 口座候補 {jobPreview.accounts.length} 件</p>
                                <p className="text-xs text-yellow-800 mt-1">同じ統合キーを設定した口座は1口座として登録できます。</p>
                            </div>
                            <div className="flex flex-col gap-2 md:flex-row md:items-center">
                                <div className="flex flex-col">
                                    <label className="text-xs text-yellow-900">既存案件を選択</label>
                                    <select
                                        value={newCaseName ? '' : (selectedCaseId || '')}
                                        onChange={(e) => {
                                            setNewCaseName('');
                                            handleCaseSelectChange(e);
                                        }}
                                        className="p-2 border rounded-md text-sm"
                                        disabled={jobImportStatus === 'applying'}
                                    >
                                        <option value="">案件を選択</option>
                                        {cases.map((item) => (
                                            <option key={item.id} value={item.id}>{item.name}</option>
                                        ))}
                                    </select>
                                </div>
                                <div className="flex flex-col">
                                    <label className="text-xs text-yellow-900">新しい案件名（任意）</label>
                                    <input
                                        type="text"
                                        value={newCaseName}
                                        onChange={(e) => setNewCaseName(e.target.value)}
                                        placeholder="例: 佐藤家_2025"
                                        className="p-2 border rounded-md text-sm"
                                        disabled={jobImportStatus === 'applying'}
                                    />
                                </div>
                            </div>
                        </div>

                        <div className="grid gap-4 md:grid-cols-2">
                            {jobPreview.accounts.map((account) => {
                                const config = jobImportMappings[account.assetId] || { mode: 'new' };
                                return (
                                    <div key={account.assetId} className="bg-white border border-yellow-200 rounded-xl p-4 space-y-3 shadow-sm">
                                        <div>
                                            <h4 className="text-lg font-semibold text-gray-900">{account.accountName || '預金口座'}</h4>
                                            <p className="text-xs text-gray-500">口座番号: {account.accountNumber || '不明'}</p>
                                            <p className="text-xs text-gray-500">取引件数: {account.transactionCount} ／ 入金 {account.totalDeposit} ／ 出金 {account.totalWithdrawal}</p>
                                        </div>
                                        <div className="space-y-2">
                                            <label className="flex items-center space-x-2 text-sm">
                                                <input
                                                    type="radio"
                                                    checked={config.mode === 'new'}
                                                    onChange={() => handleJobMappingChange(account.assetId, { mode: 'new', targetAccountId: null })}
                                                />
                                                <span>新規口座として登録</span>
                                            </label>
                                            {config.mode === 'new' && (
                                                <div className="grid grid-cols-1 gap-2 text-sm">
                                                    <input
                                                        type="text"
                                                        value={config.accountName ?? account.accountName ?? ''}
                                                        onChange={(e) => handleJobMappingChange(account.assetId, { accountName: e.target.value })}
                                                        placeholder="口座名"
                                                        className="p-2 border rounded-md"
                                                    />
                                                    <input
                                                        type="text"
                                                        value={config.accountNumber ?? account.accountNumber ?? ''}
                                                        onChange={(e) => handleJobMappingChange(account.assetId, { accountNumber: e.target.value })}
                                                        placeholder="口座番号"
                                                        className="p-2 border rounded-md"
                                                    />
                                                </div>
                                            )}
                                            <label className="flex items-center space-x-2 text-sm">
                                                <input
                                                    type="radio"
                                                    checked={config.mode === 'group'}
                                                    onChange={() => handleJobMappingChange(account.assetId, { mode: 'group', mergeGroupKey: config.mergeGroupKey || account.assetId })}
                                                />
                                                <span>他の新規口座と統合</span>
                                            </label>
                                            {config.mode === 'group' && (
                                                <div className="grid grid-cols-1 gap-2 text-sm">
                                                    <input
                                                        type="text"
                                                        value={config.mergeGroupKey || ''}
                                                        onChange={(e) => handleJobMappingChange(account.assetId, { mergeGroupKey: e.target.value })}
                                                        placeholder="統合グループキー (同じ値でまとめる)"
                                                        className="p-2 border rounded-md"
                                                    />
                                                    <input
                                                        type="text"
                                                        value={config.mergeGroupName ?? config.accountName ?? account.accountName ?? ''}
                                                        onChange={(e) => handleJobMappingChange(account.assetId, { mergeGroupName: e.target.value })}
                                                        placeholder="統合後の口座名"
                                                        className="p-2 border rounded-md"
                                                    />
                                                    <input
                                                        type="text"
                                                        value={config.mergeGroupNumber ?? config.accountNumber ?? ''}
                                                        onChange={(e) => handleJobMappingChange(account.assetId, { mergeGroupNumber: e.target.value })}
                                                        placeholder="統合後の口座番号"
                                                        className="p-2 border rounded-md"
                                                    />
                                                </div>
                                            )}
                                            <label className="flex items-center space-x-2 text-sm">
                                                <input
                                                    type="radio"
                                                    checked={config.mode === 'merge'}
                                                    onChange={() => handleJobMappingChange(account.assetId, { mode: 'merge' })}
                                                />
                                                <span>既存口座とマージ</span>
                                            </label>
                                            {config.mode === 'merge' && (
                                                <select
                                                    value={config.targetAccountId || ''}
                                                    onChange={(e) => handleJobMappingChange(account.assetId, { targetAccountId: e.target.value })}
                                                    className="p-2 border rounded-md text-sm"
                                                >
                                                    <option value="">既存口座を選択</option>
                                                    {accounts.map((acc) => (
                                                        <option key={acc.id} value={acc.id}>{acc.name} / {acc.number || '番号なし'}</option>
                                                    ))}
                                                </select>
                                            )}
                                        </div>
                                    </div>
                                );
                            })}
                        </div>
                        {jobImportError && <p className="text-sm text-red-600 bg-red-100 border border-red-200 rounded-lg p-2">{jobImportError}</p>}
                        <div className="flex justify-end">
                            <MainButton
                                onClick={handleApplyJobImport}
                                Icon={jobImportStatus === 'applying' ? Loader2 : Save}
                                className="bg-yellow-600 hover:bg-yellow-700"
                                disabled={jobImportStatus === 'applying'}
                            >
                                {jobImportStatus === 'applying' ? '取り込み中…' : 'この内容で案件に反映'}
                            </MainButton>
                        </div>
                    </section>
                )}
                {/* タブナビゲーション */}
                <div className="flex border-b border-gray-200 overflow-x-auto whitespace-nowrap">
                    {/* 口座登録/ホームタブ (機能強化済み) */}
                    <TabButton 
                        tabId="register" 
                        label="設定・並べ替え" 
                        Icon={CreditCard} 
                    />

                    {/* 統合タブ */}
                    <TabButton 
                        tabId="integrated" 
                        label="統合一覧" 
                        Icon={ArrowDownUp} 
                        className="bg-purple-100 hover:bg-purple-200"
                        badge={totalTransactionCount}
                    />

                    {/* 口座ごとのタブ (口座番号を追加) */}
                    {accounts.map(account => (
                        <TabButton
                            key={account.id}
                            tabId={account.id}
                            label={account.name}
                            Icon={List}
                            badge={transactions.filter(t => t.accountId === account.id).length}
                            number={account.number} // 口座番号を渡す
                        />
                    ))}
                </div>

                {/* メインコンテンツ */}
                <main className="bg-white rounded-b-xl shadow-xl min-h-[60vh]">
                    {renderContent()}
                </main>
            </div>

            {/* 新規口座追加モーダル (機能分離) */}
            <AddAccountModal 
                isOpen={showAddAccountModal}
                onClose={() => setShowAddAccountModal(false)}
                onCreateAccount={handleCreateAccount}
                caseName={cases.find((item) => item.id === selectedCaseId)?.name}
            />

            <EditAccountModal
                isOpen={!!editingAccount}
                onClose={() => setEditingAccount(null)}
                account={editingAccount}
                onUpdateAccount={handleUpdateAccount}
            />
            
            {/* 取引編集モーダル */}
            <EditTransactionModal
                isOpen={!!editingTransaction}
                onClose={() => setEditingTransaction(null)}
                transaction={editingTransaction}
                onUpdateTransaction={handleUpdateTransaction}
            />

            {/* データ管理モーダル */}
            <ExportModal
                isOpen={showExportModal}
                onClose={() => setShowExportModal(false)}
                accounts={accounts}
                transactions={transactions}
            />
            <ImportModal
                isOpen={showImportModal}
                onClose={() => setShowImportModal(false)}
                onImport={handleImportData}
                caseName={cases.find((item) => item.id === selectedCaseId)?.name}
            />

            <UsageGuideModal isOpen={showGuide} onClose={() => setShowGuide(false)} />

            <PendingImportModal
                isOpen={showPendingImportModal}
                onClose={() => setShowPendingImportModal(false)}
                pendingImports={pendingImports}
                caseName={cases.find((item) => item.id === selectedCaseId)?.name}
                onApply={(entry, overrides = {}) => handleImportPendingEntry(entry, { targetCaseId: selectedCaseId, ...overrides })}
                onManual={handleManualAddAccount}
                onDismiss={handleDismissPendingEntry}
                status={pendingImportStatus}
                error={pendingImportError}
            />

            {/* 認証状態の表示 */}
            <div className="max-w-7xl mx-auto p-4">
                <StatusMessage loading={loading && !error} error={error} userId={userId} />
            </div>
        </div>
    );
};

export default LedgerApp;
