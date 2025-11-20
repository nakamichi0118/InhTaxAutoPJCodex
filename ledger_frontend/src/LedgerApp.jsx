import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { List, Plus, Minus, CreditCard, Save, Trash2, X, Clipboard, ArrowDownUp, Edit, ChevronUp, ChevronDown, FileDown, Loader2, FileUp } from 'lucide-react';

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
  const [name, setName] = useState('');
  const [number, setNumber] = useState('');
  const [message, setMessage] = useState('');

    const handleSaveAccount = async () => {
        if (!name || !number) {
            setMessage('名義人と口座番号を入力してください。');
            return;
        }

        try {
            await onCreateAccount({ name, number });
            setMessage('口座情報が正常に登録されました！');
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
        <InputField label="名義人 (口座名)" id="accountName" value={name} onChange={(e) => setName(e.target.value)} required icon={Clipboard} placeholder="山田 太郎" />
        <InputField label="口座番号" id="accountNumber" value={number} onChange={(e) => setNumber(e.target.value.replace(/[^0-9]/g, ''))} required icon={CreditCard} placeholder="1234567" type="tel" inputMode="numeric" pattern="[0-9]*" />
        
        {message && <p className={`p-3 rounded-lg my-3 ${message.includes('エラー') ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}`}>{message}</p>}

        <MainButton type="submit" Icon={Plus} className="w-full mt-4 bg-green-600 hover:bg-green-700">
          新規口座を登録
        </MainButton>
      </form>
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
                    <div key={acc.id} className="flex justify-between items-center bg-white p-4 rounded-lg border border-gray-200 shadow-sm">
                    <div className="text-sm">
                        <p className="font-bold text-gray-800 text-lg">{acc.name}</p>
                        <p className="text-gray-500 font-mono text-xs">口座番号: {acc.number}</p>
                        <p className="text-gray-400 text-xs mt-1">表示順: {index + 1}番目</p>
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

    const accountTransactions = transactions.filter(t => t.accountId === account.id)
        .sort((a, b) => {
            const dateComparison = new Date(a.date) - new Date(b.date);
            if (dateComparison !== 0) return dateComparison;
            return a.id.localeCompare(b.id);
        });

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
                <h3 className="text-xl font-bold text-gray-700 mb-3">取引履歴 ({accountTransactions.length}件)</h3>
                <div className="bg-white rounded-xl shadow-lg overflow-hidden">
                    <TransactionTable
                        transactions={accountTransactions}
                        accounts={[{id: account.id, name: account.name}]} // Pass the current account in an array
                        onDelete={handleDeleteTransaction}
                        onEdit={setEditingTransaction}
                        onColorChange={() => {}} // Dummy function as color is not changed here
                        showAccountInfo={false} // We are in a single account view
                    />
                </div>
            </div>
        </div>
    );
};

const TransactionTable = ({ transactions, accounts, onDelete, onEdit, onColorChange, onReorder, showAccountInfo = true }) => {
    const accountMap = useMemo(() => {
        // accountsがundefinedでないことを保証してからreduceを呼び出す
        if (!accounts) return {}; 
        return accounts.reduce((map, acc) => {
            map[acc.id] = acc;
            return map;
        }, {});
    }, [accounts]);

    const ColorPicker = ({ transactionId, currentColor }) => {
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
        }

        return (
            <div className="flex items-center justify-center space-x-1">
                {colors.map(color => (
                    <button
                        key={color}
                        onClick={() => onColorChange(transactionId, color === currentColor ? null : color)}
                        className={`w-4 h-4 rounded-full border-2 transition-transform transform hover:scale-125 ${
                            currentColor === color ? getBorderClass(color) : 'border-transparent'
                        } ${getBgClass(color)}`}
                        title={`${colorTooltips[color]}でマーク`}
                    />
                ))}
            </div>
        );
    };

    const totalWithdrawal = transactions.reduce((sum, t) => sum + (t.withdrawal || 0), 0);
    const totalDeposit = transactions.reduce((sum, t) => sum + (t.deposit || 0), 0);
    const balance = totalDeposit - totalWithdrawal;

    return (
        <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                    <tr>
                        <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">日付</th>
                        {showAccountInfo && (
                            <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">取引口座</th>
                        )}
                        <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">種別</th>
                        <th className="px-3 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">出金額</th>
                        <th className="px-3 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">入金額</th>
                        <th className="px-3 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">備考/カテゴリ</th>
                        {showAccountInfo && (
                            <th className="px-3 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">着色</th>
                        )}
                        {showAccountInfo && (
                            <th className="px-1 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">順序</th>
                        )}
                        <th className="px-3 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider" colSpan="2">操作</th>
                    </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                    {transactions.length === 0 ? (
                        <tr>
                            <td colSpan={showAccountInfo ? 10 : 7} className="px-6 py-4 whitespace-nowrap text-sm text-gray-500 text-center">
                                取引履歴はありません。
                            </td>
                        </tr>
                    ) : (
                        transactions.map((t, index) => {
                            const account = accountMap[t.accountId];
                            const getRowClass = (color) => {
                                switch(color) {
                                    case 'green': return 'bg-green-50 hover:bg-green-100';
                                    case 'blue': return 'bg-blue-50 hover:bg-blue-100';
                                    case 'yellow': return 'bg-yellow-50 hover:bg-yellow-100';
                                    case 'pink': return 'bg-pink-50 hover:bg-pink-100';
                                    default: return 'hover:bg-gray-50';
                                }
                            }
                            const rowClass = getRowClass(t.rowColor);
                            return (
                                <tr key={t.id} className={`${rowClass} transition duration-150`}>
                                    <td className="px-3 py-3 whitespace-nowrap text-sm font-medium text-gray-900">{t.date}</td>
                                    {showAccountInfo && (
                                        <td className="px-3 py-3 whitespace-nowrap text-xs text-gray-500">
                                            <p className="font-semibold text-gray-800">{account?.name || '不明な口座'}</p>
                                            <p className="font-mono text-gray-400">No. {account?.number || '---'}</p>
                                        </td>
                                    )}
                                    <td className="px-3 py-3 whitespace-nowrap text-sm text-gray-700">{t.type || '---'}</td>
                                    <td className="px-3 py-3 whitespace-nowrap text-sm text-right text-red-600 font-mono">
                                        {(t.withdrawal || 0) > 0 ? formatCurrency(t.withdrawal) : '-'}
                                    </td>
                                    <td className="px-3 py-3 whitespace-nowrap text-sm text-right text-green-600 font-mono">
                                        {(t.deposit || 0) > 0 ? formatCurrency(t.deposit) : '-'}
                                    </td>
                                    <td className="px-3 py-3 max-w-xs overflow-hidden text-ellipsis text-sm text-gray-700">
                                        <p>{t.memo || '---'}</p>
                                    </td>
                                    
                                    {showAccountInfo && (
                                        <td className="px-3 py-3 whitespace-nowrap text-center">
                                            <ColorPicker transactionId={t.id} currentColor={t.rowColor} />
                                        </td>
                                    )}

                                    {showAccountInfo && (
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
                                            onClick={() => onEdit({...t, accountName: account?.name})}
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
                {/* 合計フッター */}
                <tfoot className="bg-gray-100 font-bold">
                    <tr>
                        <td colSpan={showAccountInfo ? 3 : 2} className="px-3 py-3 text-left text-base text-gray-800">合計</td>
                        <td className="px-3 py-3 text-right text-red-600 text-base font-mono">{formatCurrency(totalWithdrawal)}</td>
                        <td className="px-3 py-3 text-right text-green-600 text-base font-mono">{formatCurrency(totalDeposit)}</td>
                        <td className="px-3 py-3 text-right text-base text-gray-800" colSpan={showAccountInfo ? 5 : 3}>
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
}) => {
    const [message, setMessage] = useState(''); // メッセージ表示用ステートを追加

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

    const handleReorderTransaction = async (movedItem, adjacentItem) => {
        try {
            const movedIndex = integratedTransactions.findIndex(t => t.id === movedItem.id);
            const adjacentIndex = integratedTransactions.findIndex(t => t.id === adjacentItem.id);

            const isMovingDown = movedIndex < adjacentIndex;

            let prevItem, nextItem;

            if (isMovingDown) {
                prevItem = adjacentItem;
                nextItem = integratedTransactions[adjacentIndex + 1];
            } else {
                prevItem = integratedTransactions[adjacentIndex - 1];
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
                const reorderedTransactions = [...integratedTransactions];
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
            <p className="text-gray-600">
                すべての口座の取引が集約されています。「順序」列の矢印ボタンで表示順を変更できます。
            </p>
            
            {/* メッセージ表示 */}
            {message && <p className={`p-3 rounded-lg my-3 text-sm flex items-center justify-center ${message.includes('エラー') ? 'bg-red-100 text-red-700' : 'bg-blue-100 text-blue-700'}`}><Loader2 size={16} className="animate-spin mr-2" /> {message}</p>}

            <div className="bg-white rounded-xl shadow-lg overflow-hidden">
                <TransactionTable
                    transactions={integratedTransactions}
                    accounts={allAccounts}
                    onDelete={handleDeleteTransaction}
                    onEdit={setEditingTransaction}
                    onReorder={handleReorderTransaction}
                    onColorChange={handleColorChange}
                    showAccountInfo={true}
                />
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
        async ({ name, number }) => {
            if (!selectedCaseId) {
                throw new Error('案件が選択されていません。');
            }
            await callLedgerApi('/accounts', {
                method: 'POST',
                body: { name, number, caseId: selectedCaseId },
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
                await callLedgerApi('/import', {
                    method: 'POST',
                    body: {
                        caseId: caseIdToUse,
                        newCaseName: newCaseName || null,
                        accounts: payload.accounts,
                        transactions: payload.transactions,
                    },
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
                setPendingImportError(err.message);
            }
        },
        [callLedgerApi, convertAssetsToLedgerPayload, removePendingImportEntry, fetchCases, refreshState, selectedCaseId],
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
            const name = asset?.asset_name || (Array.isArray(asset?.owner_name) && asset.owner_name[0]) || `口座${index + 1}`;
            accounts.push({
                id: accountId,
                name,
                number: identifiers.primary || identifiers.secondary || '',
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
                    if (config.mode === 'merge' && !config.targetAccountId) {
                        throw new Error(`${account.accountName || '口座'} のマージ先を選択してください。`);
                    }
                    return {
                        assetId: account.assetId,
                        mode: config.mode || 'new',
                        targetAccountId: config.targetAccountId || null,
                        accountName: config.accountName || account.accountName,
                        accountNumber: config.accountNumber || account.accountNumber,
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
        <div className="min-h-screen bg-gray-50 font-sans">
            <header className="bg-white shadow-md">
                <div className="max-w-7xl mx-auto py-4 px-4 sm:px-6 lg:px-8 flex justify-between items-center">
                    <h1 className="text-3xl font-extrabold text-blue-800 flex items-center space-x-2">
                        <List size={30} className="text-blue-500" />
                        <span>入出金検討表作成ツール</span>
                    </h1>
                    {/* 右上のボタンを新規口座登録に変更 */}
                    <MainButton Icon={Plus} onClick={() => setShowAddAccountModal(true)} className="bg-green-600 hover:bg-green-700 px-4 py-2">
                        新規口座を登録
                    </MainButton>
                </div>
            </header>

            <div className="max-w-7xl mx-auto mt-4 px-4 sm:px-6 lg:px-8">
                <div className="mb-4 text-xs text-gray-600 bg-yellow-50 border border-yellow-200 rounded-md p-3 leading-relaxed">
                    この画面では Railway 上の FastAPI + Ledger API にデータを保存します。ブラウザごとに匿名IDが割り当てられるため、端末を変えた場合は別ID扱いになります。
                    共有したい場合は JSON でエクスポートし、必要に応じてインポートしてください。
                </div>
                <div className="flex flex-wrap items-end gap-4 bg-white border border-gray-200 rounded-xl p-4 shadow-sm mb-4">
                    <div className="flex flex-col">
                        <label className="text-sm text-gray-600 mb-1">案件を選択</label>
                        <select
                            value={selectedCaseId || ''}
                            onChange={handleCaseSelectChange}
                            className="p-2 border rounded-lg min-w-[220px]"
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

                {jobPreview && (
                    <section className="bg-yellow-50 border border-yellow-200 rounded-2xl p-5 mb-5 space-y-4">
                        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
                            <div>
                                <h3 className="text-xl font-semibold text-yellow-900">OCR結果を案件へ取り込み</h3>
                                <p className="text-sm text-yellow-800">ジョブID: {jobPreview.jobId} ／ 口座候補 {jobPreview.accounts.length} 件</p>
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
