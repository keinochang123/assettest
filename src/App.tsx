import React, { useState, useEffect, useRef } from 'react';
import { 
  collection, 
  doc, 
  onSnapshot, 
  query, 
  where, 
  addDoc, 
  updateDoc, 
  Timestamp, 
  orderBy,
  deleteDoc,
  getDocs,
  writeBatch
} from 'firebase/firestore';
import { 
  LayoutDashboard, 
  ClipboardList, 
  ScanQrCode, 
  FileSpreadsheet, 
  Plus, 
  Search, 
  CheckCircle2, 
  AlertCircle, 
  History,
  Download,
  Upload,
  Camera,
  X,
  RefreshCw,
  User as UserIcon,
  Building2,
  Tag,
  Trash2
} from 'lucide-react';
import { 
  PieChart, 
  Pie, 
  Cell, 
  ResponsiveContainer, 
  Legend, 
  Tooltip as RechartsTooltip 
} from 'recharts';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { Html5QrcodeScanner, Html5Qrcode } from 'html5-qrcode';
import { motion, AnimatePresence } from 'motion/react';

import { db } from './firebase';
import { cn } from './lib/utils';
import { 
  InventoryPlan, 
  Asset, 
  UserProfile, 
  CheckResult 
} from './types';

// --- Components ---

const Button = ({ 
  children, 
  className, 
  variant = 'primary', 
  ...props 
}: React.ButtonHTMLAttributes<HTMLButtonElement> & { variant?: 'primary' | 'secondary' | 'outline' | 'danger' | 'ghost' }) => {
  const variants = {
    primary: 'bg-blue-600 text-white hover:bg-blue-700',
    secondary: 'bg-slate-800 text-white hover:bg-slate-900',
    outline: 'border border-slate-200 bg-white hover:bg-slate-50 text-slate-700',
    danger: 'bg-red-500 text-white hover:bg-red-600',
    ghost: 'hover:bg-slate-100 text-slate-600',
  };

  return (
    <button 
      className={cn(
        'px-4 py-2 rounded-lg font-medium transition-all flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed',
        variants[variant],
        className
      )}
      {...props}
    >
      {children}
    </button>
  );
};

const Card = ({ children, className, ...props }: { children: React.ReactNode; className?: string } & React.HTMLAttributes<HTMLDivElement>) => (
  <div className={cn('bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden', className)} {...props}>
    {children}
  </div>
);

const Badge = ({ children, variant = 'neutral' }: { children: React.ReactNode; variant?: 'neutral' | 'success' | 'warning' | 'danger' | 'info' }) => {
  const variants = {
    neutral: 'bg-slate-100 text-slate-600',
    success: 'bg-emerald-100 text-emerald-700',
    warning: 'bg-amber-100 text-amber-700',
    danger: 'bg-rose-100 text-rose-700',
    info: 'bg-blue-100 text-blue-700',
  };
  return (
    <span className={cn('px-2 py-0.5 rounded-full text-xs font-semibold', variants[variant])}>
      {children}
    </span>
  );
};

const Input = ({ label, ...props }: { label?: string } & React.InputHTMLAttributes<HTMLInputElement>) => (
  <div className="space-y-1">
    {label && <label className="block text-sm font-medium text-slate-700">{label}</label>}
    <input 
      className="w-full px-3 py-2 rounded-lg border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all"
      {...props}
    />
  </div>
);

// --- Utilities ---

function withTimeout<T>(promise: Promise<T>, timeoutMs: number = 5500): Promise<T> {
  return Promise.race([
    promise,
    new Promise<T>((_, reject) => 
      setTimeout(() => {
        const err = new Error('Firebase operation timeout (Quota limit exceeded or offline)');
        (err as any).code = 'timeout';
        reject(err);
      }, timeoutMs)
    )
  ]);
}

const formatExcelDate = (val: any): string => {
  if (val === undefined || val === null || val === '') return '';
  
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
  }
  
  const strVal = String(val).trim();
  // Check if it's a numeric string (Excel serial number)
  if (/^\d+(\.\d+)?$/.test(strVal)) {
    const num = Number(strVal);
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + num * 86400000);
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
  }
  
  return strVal;
};

const formatCreatedAt = (createdAt: any): string => {
  if (!createdAt) return '';
  if (typeof createdAt.toDate === 'function') {
    return createdAt.toDate().toLocaleDateString();
  }
  if (createdAt.seconds) {
    return new Date(createdAt.seconds * 1000).toLocaleDateString();
  }
  if (typeof createdAt === 'string' || typeof createdAt === 'number') {
    return new Date(createdAt).toLocaleDateString();
  }
  return '';
};

// --- Main App ---

export default function App() {
  const [activeTab, setActiveTab] = useState<'plans' | 'pda' | 'reports'>('plans');
  const [isLocalMode, setIsLocalMode] = useState<boolean>(() => {
    return localStorage.getItem('is_local_mode') === 'true';
  });

  useEffect(() => {
    localStorage.setItem('is_local_mode', String(isLocalMode));
  }, [isLocalMode]);
  
  // Plans State
  const [plans, setPlans] = useState<InventoryPlan[]>([]);
  const [selectedPlanId, setSelectedPlanId] = useState<string | null>(null);
  const [showNewPlanModal, setShowNewPlanModal] = useState(false);
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [newPlan, setNewPlan] = useState({ name: '', description: '', scope: '' });
  const [importFile, setImportFile] = useState<File | null>(null);
  const [isCreatingPlan, setIsCreatingPlan] = useState(false);
  const [isClearing, setIsClearing] = useState(false);
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'error' | 'info' } | null>(null);

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'info') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 3000);
  };

  // Assets State
  const [assets, setAssets] = useState<Asset[]>([]);
  const [batchAssets, setBatchAssets] = useState<Asset[]>([]);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [scanning, setScanning] = useState(false);
  const [searchCode, setSearchCode] = useState('');
  const [pdaRemarks, setPdaRemarks] = useState('');
  const [pdaStatus, setPdaStatus] = useState<CheckResult>('normal');
  const [pdaCustodian, setPdaCustodian] = useState('');
  const [pdaLocation, setPdaLocation] = useState('');
  const [pdaOffice, setPdaOffice] = useState('');
  const [pdaError, setPdaError] = useState<string | null>(null);

  const scannerRef = useRef<Html5QrcodeScanner | null>(null);

  // Plans Listener
  useEffect(() => {
    const clearAllTimestampStr = localStorage.getItem('clear_all_timestamp');
    const clearAllTime = clearAllTimestampStr ? Number(clearAllTimestampStr) : 0;

    if (isLocalMode) {
      const localPlansStr = localStorage.getItem('local_plans');
      let localPlans: InventoryPlan[] = [];
      try {
        localPlans = localPlansStr ? JSON.parse(localPlansStr) : [];
      } catch (e) {
        console.error("Failed to parse local plans", e);
      }
      
      const clearedPlanIdsStr = localStorage.getItem('cleared_plan_ids') || '[]';
      let clearedPlanIds: string[] = [];
      try { clearedPlanIds = JSON.parse(clearedPlanIdsStr); } catch (e) {}

      const filtered = localPlans.filter(p => {
        if (!p.id) return false;
        if (clearedPlanIds.includes(p.id)) return false;
        if (p.createdAt) {
          const createdAtAny = p.createdAt as any;
          const createdAtMs = typeof createdAtAny.toMillis === 'function'
            ? createdAtAny.toMillis()
            : (typeof createdAtAny.toDate === 'function' ? createdAtAny.toDate().getTime() : new Date(createdAtAny).getTime());
          if (createdAtMs <= clearAllTime) return false;
        }
        return true;
      });
      setPlans(filtered);
      return;
    }

    const q = query(collection(db, 'plans'), orderBy('createdAt', 'desc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const clearedPlanIdsStr = localStorage.getItem('cleared_plan_ids') || '[]';
      let clearedPlanIds: string[] = [];
      try { clearedPlanIds = JSON.parse(clearedPlanIdsStr); } catch (e) {}

      const plansData = snapshot.docs
        .map(doc => ({ id: doc.id, ...doc.data() } as InventoryPlan))
        .filter(p => {
          if (!p.id) return false;
          if (clearedPlanIds.includes(p.id)) return false;
          if (p.createdAt) {
            const createdAtAny = p.createdAt as any;
            const createdAtMs = typeof createdAtAny.toMillis === 'function'
              ? createdAtAny.toMillis()
              : (typeof createdAtAny.toDate === 'function' ? createdAtAny.toDate().getTime() : new Date(createdAtAny).getTime());
            if (createdAtMs <= clearAllTime) return false;
          }
          return true;
        });
      setPlans(plansData);
    }, (error: any) => {
      console.error("Firestore plans listener failed, auto-switching to local mode:", error);
      setIsLocalMode(true);
      showToast('已切換至「本機模擬儲存模式」（雲端額度已滿或無權限）', 'info');
    });
    return unsubscribe;
  }, [isLocalMode]);

  // Synchronize selectedPlanId with available plans when plans list changes or mode switches
  useEffect(() => {
    if (plans.length > 0) {
      const exists = plans.some(p => p.id === selectedPlanId);
      if (!exists) {
        setSelectedPlanId(plans[0].id || null);
      }
    } else {
      setSelectedPlanId(null);
    }
  }, [plans, isLocalMode]);

  // Assets Listener
  useEffect(() => {
    if (!selectedPlanId) {
      setAssets([]);
      return;
    }

    if (isLocalMode) {
      const localAssetsStr = localStorage.getItem('local_assets');
      const allLocalAssets: Asset[] = localAssetsStr ? JSON.parse(localAssetsStr) : [];
      
      const clearedAssetIdsStr = localStorage.getItem('cleared_asset_ids') || '[]';
      let clearedAssetIds: string[] = [];
      try { clearedAssetIds = JSON.parse(clearedAssetIdsStr); } catch (e) {}

      const filteredAssets = allLocalAssets
        .filter(a => a.planId === selectedPlanId)
        .filter(a => {
          if (!a.id) return false;
          if (clearedAssetIds.includes(a.id)) return false;
          return true;
        });
      setAssets(filteredAssets);
      return;
    }

    const q = query(collection(db, 'assets'), where('planId', '==', selectedPlanId));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const clearedAssetIdsStr = localStorage.getItem('cleared_asset_ids') || '[]';
      let clearedAssetIds: string[] = [];
      try { clearedAssetIds = JSON.parse(clearedAssetIdsStr); } catch (e) {}

      const assetsData = snapshot.docs
        .map(doc => ({ id: doc.id, ...doc.data() } as Asset))
        .filter(a => {
          if (!a.id) return false;
          if (clearedAssetIds.includes(a.id)) return false;
          return true;
        });
      setAssets(assetsData);
    }, (error: any) => {
      console.error("Firestore assets listener failed, auto-switching to local mode:", error);
      setIsLocalMode(true);
      showToast('已切換至「本機模擬儲存模式」（雲端額度已滿或無權限）', 'info');
    });
    return unsubscribe;
  }, [selectedPlanId, isLocalMode]);

  // PDA Scanner Effect
  useEffect(() => {
    let html5QrCode: Html5Qrcode | null = null;
    
    if (scanning) {
      const startScanner = async () => {
        try {
          // Ensure the element exists
          const element = document.getElementById("reader");
          if (!element) return;

          html5QrCode = new Html5Qrcode("reader");
          scannerRef.current = html5QrCode;

          await html5QrCode.start(
            { facingMode: "environment" }, // Force back camera
            {
              fps: 10,
              qrbox: { width: 250, height: 250 },
            },
            (decodedText) => {
              onScanSuccess(decodedText);
              // Stop after success
              if (html5QrCode) {
                html5QrCode.stop().catch(console.error);
                setScanning(false);
              }
            },
            (errorMessage) => {
              // Ignore common "no code found" errors
              if (errorMessage.includes("NotFoundException")) return;
              onScanFailure(errorMessage);
            }
          );
        } catch (err: any) {
          console.error("Scanner start failed", err);
          let msg = "相機啟動失敗。";
          if (err?.name === 'NotAllowedError' || err?.message?.includes('Permission denied')) {
            msg = "相機權限遭拒。請在瀏覽器設定中允許相機權限。";
          }
          if (window.self !== window.top) {
            msg += " 偵測到您在內嵌視窗中操作，請務必點擊右上角「在新分頁開啟」圖示以獲得相機權限。";
          }
          setPdaError(msg);
          setScanning(false);
        }
      };
      
      const timer = setTimeout(startScanner, 500);
      return () => {
        clearTimeout(timer);
        if (html5QrCode && html5QrCode.isScanning) {
          html5QrCode.stop().catch(console.error);
        }
      };
    }
  }, [scanning]);

  const onScanSuccess = (decodedText: string) => {
    setScanning(false);
    handleSearch(decodedText);
  };

  const onScanFailure = (error: any) => {
    // Some browsers block camera in iframes. If we detect a permission error, show a helpful message.
    if (error?.name === 'NotAllowedError' || error?.message?.includes('Permission denied')) {
      setPdaError("相機權限遭拒。請確認瀏覽器設定，或點擊右上角「在新分頁開啟」以正常運作。");
      setScanning(false);
    }
  };

  const handleSearch = (code: string) => {
    const searchVal = code.trim();
    if (!searchVal) return;
    setPdaError(null);
    
    // Case-insensitive and trimmed search
    const asset = assets.find(a => 
      (a.assetCode || '').trim().toLowerCase() === searchVal.toLowerCase()
    );
    
    if (asset) {
      // Check if already in batch
      const existingIndex = batchAssets.findIndex(a => 
        (a.assetCode || '').trim().toLowerCase() === searchVal.toLowerCase()
      );
      if (existingIndex >= 0) {
        setEditingIndex(existingIndex);
        const item = batchAssets[existingIndex];
        setPdaRemarks(item.checkRemark || '');
        setPdaStatus(item.checkResult || 'normal');
        setPdaCustodian(item.updatedCustodian || '');
        setPdaLocation(item.updatedLocation || '');
        setPdaOffice(item.updatedOffice || '');
      } else {
        const newBatch = [...batchAssets, asset];
        setBatchAssets(newBatch);
        setEditingIndex(newBatch.length - 1);
        setPdaRemarks(asset.checkRemark || '');
        setPdaStatus(asset.checkResult || 'normal');
        setPdaCustodian(asset.updatedCustodian || '');
        setPdaLocation(asset.updatedLocation || '');
        setPdaOffice(asset.updatedOffice || '');
      }
      setSearchCode('');
    } else {
      setPdaError(`找不到資產編號: ${code}`);
      setTimeout(() => setPdaError(null), 5000);
    }
  };

  const handleSaveCurrentEdit = () => {
    if (editingIndex === null) return;
    const updatedBatch = [...batchAssets];
    updatedBatch[editingIndex] = {
      ...updatedBatch[editingIndex],
      checkResult: pdaStatus,
      checkRemark: pdaRemarks,
      updatedCustodian: pdaCustodian,
      updatedLocation: pdaLocation,
      updatedOffice: pdaOffice,
      status: 'checked'
    };
    setBatchAssets(updatedBatch);
    setEditingIndex(null);
  };

  const handleSubmitBatch = async () => {
    if (batchAssets.length === 0) return;
    try {
      if (isLocalMode) {
        const localAssetsStr = localStorage.getItem('local_assets');
        const allLocalAssets: Asset[] = localAssetsStr ? JSON.parse(localAssetsStr) : [];
        
        batchAssets.forEach(batchAsset => {
          const idx = allLocalAssets.findIndex(a => a.id === batchAsset.id);
          if (idx >= 0) {
            allLocalAssets[idx] = {
              ...allLocalAssets[idx],
              status: 'checked',
              checkResult: batchAsset.checkResult || 'normal',
              checkRemark: batchAsset.checkRemark || '',
              checkTime: Timestamp.now(),
              checkBy: 'system',
              updatedCustodian: batchAsset.updatedCustodian || '',
              updatedLocation: batchAsset.updatedLocation || '',
              updatedOffice: batchAsset.updatedOffice || ''
            };
          }
        });
        
        localStorage.setItem('local_assets', JSON.stringify(allLocalAssets));
        
        // Refresh local assets list for current selectedPlanId
        const filtered = allLocalAssets.filter(a => a.planId === selectedPlanId);
        setAssets(filtered);
        
        showToast('整批盤點結果已提交成功！(儲存至本機)', 'success');
        setBatchAssets([]);
        setEditingIndex(null);
        return;
      }

      // Firebase Mode
      const BATCH_SIZE = 450;
      const batches = [];
      let currentBatch = writeBatch(db);
      let count = 0;

      batchAssets.forEach((asset, index) => {
        const assetRef = doc(db, 'assets', asset.id!);
        currentBatch.update(assetRef, {
          status: 'checked',
          checkResult: asset.checkResult || 'normal',
          checkRemark: asset.checkRemark || '',
          checkTime: Timestamp.now(),
          checkBy: 'system',
          updatedCustodian: asset.updatedCustodian || '',
          updatedLocation: asset.updatedLocation || '',
          updatedOffice: asset.updatedOffice || ''
        });

        count++;
        if (count === BATCH_SIZE || index === batchAssets.length - 1) {
          batches.push(currentBatch);
          currentBatch = writeBatch(db);
          count = 0;
        }
      });

      // Commit sequentially with a generous 15-second timeout per batch
      for (const batch of batches) {
        await withTimeout(batch.commit(), 15000);
      }
      
      showToast('整批盤點結果已提交成功！', 'success');
      setBatchAssets([]);
      setEditingIndex(null);
    } catch (error: any) {
      console.error('Failed to submit batch check', error);
      if (error?.code === 'resource-exhausted' || error?.code === 'timeout' || error?.message?.includes('permission-denied') || error?.message?.includes('permission') || error?.message?.includes('timeout')) {
        setIsLocalMode(true);
        showToast('雲端流量已達上限或連線逾時！已為您自動切換至「本機模擬儲存模式」。請再次點擊提交。', 'error');
      } else {
        showToast('提交失敗，請檢查網路連線', 'error');
      }
    }
  };

  const handleCreatePlan = async () => {
    if (!newPlan.name) {
      showToast('請輸入計畫名稱', 'error');
      return;
    }
    if (isCreatingPlan) return;
    
    setIsCreatingPlan(true);
    try {
      if (isLocalMode) {
        const newPlanId = 'plan_' + Date.now();
        const planObj: InventoryPlan = {
          id: newPlanId,
          name: newPlan.name,
          description: newPlan.description,
          scope: newPlan.scope,
          status: 'active',
          createdAt: Timestamp.now()
        };

        const localPlansStr = localStorage.getItem('local_plans');
        const localPlans: InventoryPlan[] = localPlansStr ? JSON.parse(localPlansStr) : [];
        localPlans.unshift(planObj);
        localStorage.setItem('local_plans', JSON.stringify(localPlans));
        setPlans(localPlans);

        if (importFile) {
          await handleImportExcel(importFile, newPlanId);
        }

        setShowNewPlanModal(false);
        setNewPlan({ name: '', description: '', scope: '' });
        setImportFile(null);
        setSelectedPlanId(newPlanId);
        showToast('本機計畫建立成功！', 'success');
        return;
      }

      // Firebase Mode
      const planRef = await withTimeout(addDoc(collection(db, 'plans'), {
        ...newPlan,
        status: 'active',
        createdAt: Timestamp.now()
      }), 15000);
      
      if (importFile) {
        await handleImportExcel(importFile, planRef.id);
      }
      
      setShowNewPlanModal(false);
      setNewPlan({ name: '', description: '', scope: '' });
      setImportFile(null);
      setSelectedPlanId(planRef.id);
      showToast('計畫建立成功！', 'success');
    } catch (error: any) {
      console.error('Failed to create plan', error);
      if (error?.code === 'resource-exhausted' || error?.code === 'timeout' || error?.message?.includes('permission-denied') || error?.message?.includes('permission') || error?.message?.includes('timeout')) {
        setIsLocalMode(true);
        showToast('雲端流量已達上限或連線逾時！已自動切換至「本機儲存模式」，請重新點擊建立計畫。', 'error');
      } else {
        showToast('建立計畫失敗，請檢查網路連線', 'error');
      }
    } finally {
      setIsCreatingPlan(false);
    }
  };

  const handleImportExcel = async (file: File, planId: string) => {
    return new Promise<void>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        try {
          const dataArray = evt.target?.result;
          const wb = XLSX.read(dataArray, { type: 'array', cellDates: true });
          const wsname = wb.SheetNames[0];
          const ws = wb.Sheets[wsname];
          const data = XLSX.utils.sheet_to_json(ws) as any[];

          const getVal = (row: any, target: string) => {
            const normalizedTarget = target.replace(/\s+/g, '');
            const key = Object.keys(row).find(k => k.replace(/\s+/g, '') === normalizedTarget);
            return key ? row[key] : undefined;
          };

          if (isLocalMode) {
            const localAssets: Asset[] = [];
            data.forEach((row) => {
              const assetNum = String(getVal(row, '資產編號') || '');
              const subNum = String(getVal(row, '子編號') || '');
              const fullAssetCode = assetNum && subNum ? `${assetNum}-${subNum}` : (assetNum || subNum);

              localAssets.push({
                id: 'asset_' + Math.random().toString(36).substring(2, 11) + '_' + Date.now(),
                planId: planId,
                assetCode: fullAssetCode,
                companyCode: String(getVal(row, '公司代碼') || ''),
                accountName: String(getVal(row, '科目名稱') || ''),
                categoryName: String(getVal(row, '類別名稱') || ''),
                assetDescription: String(getVal(row, '資產說明') || ''),
                acquisitionDate: formatExcelDate(getVal(row, '取得日期')),
                acquisitionCost: Number(getVal(row, '取得成本') || 0),
                bookValue: Number(getVal(row, '帳面價值') || 0),
                quantity: Number(getVal(row, '數量') || 0),
                unit: String(getVal(row, '單位') || ''),
                originalOffice: String(getVal(row, '室') || ''),
                originalCustodian: String(getVal(row, '保管人') || ''),
                originalLocation: String(getVal(row, '地點') || ''),
                costCenter: String(getVal(row, '成本中心') || ''),
                status: 'pending'
              });
            });

            const localAssetsStr = localStorage.getItem('local_assets');
            const allLocalAssets: Asset[] = localAssetsStr ? JSON.parse(localAssetsStr) : [];
            const merged = [...allLocalAssets, ...localAssets];
            localStorage.setItem('local_assets', JSON.stringify(merged));

            // Refresh state
            const filtered = merged.filter(a => a.planId === planId);
            setAssets(filtered);
            resolve();
            return;
          }

          // Firestore limits batches to 500 operations. We use 450 to be safe.
          const BATCH_SIZE = 450;
          const batches = [];
          let currentBatch = writeBatch(db);
          let operationCount = 0;

          data.forEach((row, index) => {
            const assetRef = doc(collection(db, 'assets'));
            
            const assetNum = String(getVal(row, '資產編號') || '');
            const subNum = String(getVal(row, '子編號') || '');
            const fullAssetCode = assetNum && subNum ? `${assetNum}-${subNum}` : (assetNum || subNum);

            currentBatch.set(assetRef, {
              planId: planId,
              assetCode: fullAssetCode,
              companyCode: String(getVal(row, '公司代碼') || ''),
              accountName: String(getVal(row, '科目名稱') || ''),
              categoryName: String(getVal(row, '類別名稱') || ''),
              assetDescription: String(getVal(row, '資產說明') || ''),
              acquisitionDate: formatExcelDate(getVal(row, '取得日期')),
              acquisitionCost: Number(getVal(row, '取得成本') || 0),
              bookValue: Number(getVal(row, '帳面價值') || 0),
              quantity: Number(getVal(row, '數量') || 0),
              unit: String(getVal(row, '單位') || ''),
              originalOffice: String(getVal(row, '室') || ''),
              originalCustodian: String(getVal(row, '保管人') || ''),
              originalLocation: String(getVal(row, '地點') || ''),
              costCenter: String(getVal(row, '成本中心') || ''),
              status: 'pending'
            });

            operationCount++;

            if (operationCount === BATCH_SIZE || index === data.length - 1) {
              batches.push(currentBatch);
              currentBatch = writeBatch(db);
              operationCount = 0;
            }
          });

          // Commit batches sequentially with a safe 15s timeout per batch to avoid random timeout switching to local mode
          for (const batch of batches) {
            await withTimeout(batch.commit(), 15000);
          }
          resolve();
        } catch (error: any) {
          console.error('Import failed', error);
          reject(error);
        }
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  };

  const handleClearAllData = async () => {
    if (isClearing) return;
    setIsClearing(true);
    
    // 1. Instantly record clear timestamp and local-first clean-up to guarantee immediate UI updates
    const now = Date.now();
    localStorage.setItem('clear_all_timestamp', String(now));

    const currentPlanIds = plans.map(p => p.id).filter(Boolean) as string[];
    const currentAssetIds = assets.map(a => a.id).filter(Boolean) as string[];

    const storedClearedPlansStr = localStorage.getItem('cleared_plan_ids') || '[]';
    let storedClearedPlans: string[] = [];
    try { storedClearedPlans = JSON.parse(storedClearedPlansStr); } catch (e) {}

    const storedClearedAssetsStr = localStorage.getItem('cleared_asset_ids') || '[]';
    let storedClearedAssets: string[] = [];
    try { storedClearedAssets = JSON.parse(storedClearedAssetsStr); } catch (e) {}

    let updatedClearedPlans = Array.from(new Set([...storedClearedPlans, ...currentPlanIds]));
    let updatedClearedAssets = Array.from(new Set([...storedClearedAssets, ...currentAssetIds]));

    localStorage.setItem('cleared_plan_ids', JSON.stringify(updatedClearedPlans));
    localStorage.setItem('cleared_asset_ids', JSON.stringify(updatedClearedAssets));

    // Reset local cache files
    localStorage.setItem('local_plans', JSON.stringify([]));
    localStorage.setItem('local_assets', JSON.stringify([]));
    
    // Reset React state variables instantly so there is NO waiting spinner or delay
    setPlans([]);
    setAssets([]);
    setBatchAssets([]);
    setEditingIndex(null);
    setSelectedPlanId(null);
    
    // Close confirmation dialog and stop the clearing spinner instantly
    setShowClearConfirm(false);
    setIsClearing(false);
    showToast('所有資料已成功清除！', 'success');

    // 2. Perform background cloud deletion asynchronously (completely out-of-band/non-blocking)
    setTimeout(async () => {
      try {
        const BATCH_SIZE = 450;
        
        // Fetch and process assets
        const assetsSnap = await withTimeout(getDocs(collection(db, 'assets')), 10000).catch(err => {
          console.warn("Background assets fetch failed (safe fallback):", err);
          return null;
        });

        if (assetsSnap && assetsSnap.size > 0) {
          const cloudAssetIds = assetsSnap.docs.map(doc => doc.id);
          const currentClearedAssetsStr = localStorage.getItem('cleared_asset_ids') || '[]';
          let currentClearedAssets: string[] = [];
          try { currentClearedAssets = JSON.parse(currentClearedAssetsStr); } catch (e) {}
          localStorage.setItem('cleared_asset_ids', JSON.stringify(Array.from(new Set([...currentClearedAssets, ...cloudAssetIds]))));

          const batches = [];
          let currentBatch = writeBatch(db);
          let count = 0;
          
          assetsSnap.docs.forEach((doc, index) => {
            currentBatch.delete(doc.ref);
            count++;
            if (count === BATCH_SIZE || index === assetsSnap.docs.length - 1) {
              batches.push(currentBatch);
              currentBatch = writeBatch(db);
              count = 0;
            }
          });

          for (const batch of batches) {
            await withTimeout(batch.commit(), 10000).catch(err => {
              console.warn("Background batch asset deletion failed (expected if quota-exhausted):", err);
            });
          }
        }

        // Fetch and process plans
        const plansSnap = await withTimeout(getDocs(collection(db, 'plans')), 10000).catch(err => {
          console.warn("Background plans fetch failed (safe fallback):", err);
          return null;
        });

        if (plansSnap && plansSnap.size > 0) {
          const cloudPlanIds = plansSnap.docs.map(doc => doc.id);
          const currentClearedPlansStr = localStorage.getItem('cleared_plan_ids') || '[]';
          let currentClearedPlans: string[] = [];
          try { currentClearedPlans = JSON.parse(currentClearedPlansStr); } catch (e) {}
          localStorage.setItem('cleared_plan_ids', JSON.stringify(Array.from(new Set([...currentClearedPlans, ...cloudPlanIds]))));

          const batches = [];
          let currentBatch = writeBatch(db);
          let count = 0;
          
          plansSnap.docs.forEach((doc, index) => {
            currentBatch.delete(doc.ref);
            count++;
            if (count === BATCH_SIZE || index === plansSnap.docs.length - 1) {
              batches.push(currentBatch);
              currentBatch = writeBatch(db);
              count = 0;
            }
          });

          for (const batch of batches) {
            await withTimeout(batch.commit(), 10000).catch(err => {
              console.warn("Background batch plan deletion failed (expected if quota-exhausted):", err);
            });
          }
        }
        console.log("Background cloud data cleanup finished successfully.");
      } catch (e) {
        console.error("Async background clear failed:", e);
      }
    }, 50);
  };

  const handleExportExcel = async () => {
    if (!selectedPlanId) return;
    const plan = plans.find(p => p.id === selectedPlanId);
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('盤點差異報表');

    // Define columns
    const columns = [
      { header: '公司代碼', key: 'companyCode', width: 15 },
      { header: '資產編號', key: 'assetCode', width: 15 },
      { header: '子編號', key: 'subCode', width: 10 },
      { header: '科目名稱', key: 'accountName', width: 20 },
      { header: '類別名稱', key: 'categoryName', width: 20 },
      { header: '資產說明', key: 'assetDescription', width: 30 },
      { header: '取得日期', key: 'acquisitionDate', width: 15 },
      { header: '取得成本', key: 'acquisitionCost', width: 15 },
      { header: '帳面價值', key: 'bookValue', width: 15 },
      { header: '數量', key: 'quantity', width: 10 },
      { header: '單位', key: 'unit', width: 10 },
      { header: '室', key: 'originalOffice', width: 15 },
      { header: '保管人', key: 'originalCustodian', width: 15 },
      { header: '地點', key: 'originalLocation', width: 20 },
      { header: '成本中心', key: 'costCenter', width: 15 },
      { header: '變更後保管人', key: 'updatedCustodian', width: 15 },
      { header: '變更後地點', key: 'updatedLocation', width: 20 },
      { header: '變更後室', key: 'updatedOffice', width: 15 },
      { header: '盤點結果', key: 'checkResult', width: 15 },
      { header: '備註', key: 'checkRemark', width: 30 }
    ];

    worksheet.columns = columns;

    // Style header
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      const headerText = cell.value as string;
      let bgColor = 'FFE0E0E0'; // Default Light Gray

      // Define color groups for specific headers
      if (['室', '保管人', '地點'].includes(headerText)) {
        bgColor = 'FFC6EFCE'; // Light Green for original info
      } else if (['變更後保管人', '變更後地點', '變更後室'].includes(headerText)) {
        bgColor = 'FFBDD7EE'; // Light Blue for updated info
      }

      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: bgColor }
      };
      cell.font = { bold: true };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // Add data
    assets.forEach((a) => {
      const assetParts = (a.assetCode || '').split('-');
      const assetCodeOnly = assetParts[0] || '';
      const subCodeOnly = assetParts.length > 1 ? assetParts[1] : '';

      const rowData = {
        companyCode: a.companyCode,
        assetCode: assetCodeOnly,
        subCode: subCodeOnly,
        accountName: a.accountName,
        categoryName: a.categoryName,
        assetDescription: a.assetDescription,
        acquisitionDate: formatExcelDate(a.acquisitionDate),
        acquisitionCost: a.acquisitionCost,
        bookValue: a.bookValue,
        quantity: a.quantity,
        unit: a.unit,
        originalOffice: a.originalOffice,
        originalCustodian: a.originalCustodian,
        originalLocation: a.originalLocation,
        costCenter: a.costCenter,
        updatedCustodian: a.updatedCustodian || '',
        updatedLocation: a.updatedLocation || '',
        updatedOffice: a.updatedOffice || '',
        checkResult: a.checkResult === 'normal' ? '正常' : a.checkResult === 'abnormal' ? '異常' : '未盤點',
        checkRemark: a.checkRemark || ''
      };

      const row = worksheet.addRow(rowData);
      
      // Style "Abnormal" result in red
      const resultCell = row.getCell('checkResult');
      if (resultCell.value === '異常') {
        resultCell.font = { color: { argb: 'FFFF0000' }, bold: true };
      }

      // Add borders to all cells
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    // Generate buffer and save
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `${plan?.name || '盤點報表'}_${new Date().toLocaleDateString()}.xlsx`);
  };

  const checkedCount = assets.filter(a => a.status === 'checked').length;
  const pendingCount = assets.length - checkedCount;
  const abnormalCount = assets.filter(a => a.checkResult === 'abnormal').length;

  const chartData = [
    { name: '已盤點', value: checkedCount, color: '#10b981' },
    { name: '未盤點', value: pendingCount, color: '#e2e8f0' }
  ];

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col lg:flex-row">
      {/* Sidebar */}
      <aside className="w-full lg:w-64 bg-slate-900 text-slate-300 flex flex-col border-r border-slate-800">
        <div className="p-6 flex items-center gap-3 border-b border-slate-800">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
            <ClipboardList className="w-5 h-5 text-white" />
          </div>
          <span className="font-bold text-white tracking-tight">資產盤點系統</span>
        </div>

        <nav className="flex-1 p-4 space-y-2">
          <button 
            onClick={() => setActiveTab('plans')}
            className={cn(
              "w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all",
              activeTab === 'plans' ? "bg-blue-600 text-white shadow-lg shadow-blue-900/20" : "hover:bg-slate-800"
            )}
          >
            <LayoutDashboard className="w-5 h-5" />
            <span className="font-medium">計畫管理</span>
          </button>
          <button 
            onClick={() => setActiveTab('pda')}
            className={cn(
              "w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all",
              activeTab === 'pda' ? "bg-blue-600 text-white shadow-lg shadow-blue-900/20" : "hover:bg-slate-800"
            )}
          >
            <ScanQrCode className="w-5 h-5" />
            <span className="font-medium">PDA 掃描</span>
          </button>
          <button 
            onClick={() => setActiveTab('reports')}
            className={cn(
              "w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all",
              activeTab === 'reports' ? "bg-blue-600 text-white shadow-lg shadow-blue-900/20" : "hover:bg-slate-800"
            )}
          >
            <FileSpreadsheet className="w-5 h-5" />
            <span className="font-medium">報表中心</span>
          </button>
        </nav>

        <div className="p-4 border-t border-slate-800">
          <button 
            onClick={() => setShowClearConfirm(true)}
            className="w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all text-rose-400 hover:bg-rose-500/10"
          >
            <Trash2 className="w-5 h-5" />
            <span className="font-medium">清除所有資料</span>
          </button>
          
          <div className="flex items-center gap-3 px-4 py-3 mt-2">
            <div className="w-8 h-8 rounded-full bg-slate-700 flex items-center justify-center overflow-hidden">
              <UserIcon className="w-4 h-4" />
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-sm font-medium text-white truncate">管理員</p>
              <p className="text-xs text-slate-500 truncate">系統管理員</p>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col h-screen overflow-hidden">
        {/* Header */}
        <header className="h-16 bg-white border-b border-slate-200 px-8 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-4">
            <h2 className="text-lg font-bold text-slate-900">
              {activeTab === 'plans' ? '盤點計畫管理' : activeTab === 'pda' ? 'PDA 行動盤點' : '盤點進度與報表'}
            </h2>
            {selectedPlanId && (
              <Badge variant="info">
                {plans.find(p => p.id === selectedPlanId)?.name}
              </Badge>
            )}
          </div>
          <div className="flex items-center gap-3">
            {/* Mode Switcher */}
            <div className="flex items-center bg-slate-100 rounded-lg p-1 text-xs font-semibold select-none gap-1 border border-slate-200">
              <button
                onClick={() => {
                  setIsLocalMode(false);
                  showToast('已嘗試切換回雲端同步模式', 'info');
                }}
                className={cn(
                  "px-2 py-1 rounded-md transition-all font-bold",
                  !isLocalMode ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-800"
                )}
              >
                雲端同步
              </button>
              <button
                onClick={() => {
                  setIsLocalMode(true);
                  showToast('已切換至本機模擬儲存模式', 'info');
                }}
                className={cn(
                  "px-2 py-1 rounded-md transition-all font-bold",
                  isLocalMode ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-800"
                )}
              >
                本機模擬
              </button>
            </div>

            <select 
              className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20"
              value={selectedPlanId || ''}
              onChange={(e) => setSelectedPlanId(e.target.value)}
            >
              <option value="" disabled>選擇盤點計畫</option>
              {plans.map(p => (
                <option key={p.id} value={p.id}>{p.name}</option>
              ))}
            </select>
            {activeTab === 'plans' && (
              <Button onClick={() => setShowNewPlanModal(true)} className="px-3 py-1.5 text-sm">
                <Plus className="w-4 h-4" />
                新增計畫
              </Button>
            )}
          </div>
        </header>

        <div className="flex-1 overflow-y-auto p-6">
          {isLocalMode && (
            <div className="mb-6 p-4 bg-amber-50 border border-amber-200 rounded-xl text-amber-800 flex flex-col sm:flex-row items-start sm:items-center gap-3 shadow-sm max-w-6xl mx-auto">
              <AlertCircle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5 sm:mt-0" />
              <div className="flex-1 text-sm font-medium">
                <span className="font-bold">⚠️ 目前已啟用「瀏覽器本機儲存模式」</span>：偵測到 Firebase 雲端資料庫額度已達上限（或未設定權限），系統已自動為您切換。所有計畫建立、Excel 匯入、盤點與報表功能皆可在此瀏覽器中正常運作！
              </div>
              <button 
                onClick={() => {
                  setIsLocalMode(false);
                  showToast('已嘗試切換回雲端同步模式', 'info');
                }}
                className="text-xs bg-amber-600 hover:bg-amber-700 text-white px-3 py-1.5 rounded-lg transition-colors font-bold self-end sm:self-auto shrink-0"
              >
                嘗試切換回雲端模式
              </button>
            </div>
          )}

          <AnimatePresence mode="wait">
            {/* Tab: Plans */}
            {activeTab === 'plans' && (
              <motion.div 
                key="plans"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-6 max-w-6xl mx-auto"
              >
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  {plans.map(plan => (
                    <Card 
                      key={plan.id} 
                      className={cn(
                        "cursor-pointer transition-all hover:ring-2 hover:ring-blue-500/50",
                        selectedPlanId === plan.id ? "ring-2 ring-blue-500" : ""
                      )}
                      onClick={() => setSelectedPlanId(plan.id!)}
                    >
                      <div className="p-6 space-y-4">
                        <div className="flex justify-between items-start">
                          <div className="w-10 h-10 bg-blue-50 rounded-lg flex items-center justify-center">
                            <ClipboardList className="w-5 h-5 text-blue-600" />
                          </div>
                          <Badge variant={plan.status === 'active' ? 'success' : 'neutral'}>
                            {plan.status === 'active' ? '進行中' : '已完成'}
                          </Badge>
                        </div>
                        <div>
                          <h3 className="font-bold text-slate-900 text-lg">{plan.name}</h3>
                          <p className="text-sm text-slate-500 line-clamp-2 mt-1">{plan.description}</p>
                        </div>
                        <div className="pt-4 border-t border-slate-100 flex items-center justify-between text-xs text-slate-400">
                          <div className="flex items-center gap-1">
                            <Building2 className="w-3 h-3" />
                            {plan.scope}
                          </div>
                          <div className="flex items-center gap-1">
                            <History className="w-3 h-3" />
                            {formatCreatedAt(plan.createdAt)}
                          </div>
                        </div>
                      </div>
                    </Card>
                  ))}
                </div>
              </motion.div>
            )}

            {/* Tab: PDA */}
            {activeTab === 'pda' && (
              <motion.div 
                key="pda"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="max-w-2xl mx-auto space-y-6"
              >
                {!selectedPlanId ? (
                  <Card className="p-12 text-center space-y-4">
                    <AlertCircle className="w-12 h-12 text-amber-500 mx-auto" />
                    <div className="space-y-2">
                      <h3 className="text-xl font-bold text-slate-900">未選擇計畫</h3>
                      <p className="text-slate-500">請先從上方選單選擇一個進行中的盤點計畫</p>
                    </div>
                  </Card>
                ) : (
                  <div className="space-y-6">
                    {/* Search & Scan Section */}
                    <Card className="p-6 space-y-6">
                      <AnimatePresence>
                        {pdaError && (
                          <motion.div 
                            initial={{ opacity: 0, y: -10 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0, y: -10 }}
                            className="p-4 bg-rose-50 border border-rose-100 rounded-xl flex items-center gap-3 text-rose-600"
                          >
                            <AlertCircle className="w-5 h-5 flex-shrink-0" />
                            <p className="text-sm font-bold">{pdaError}</p>
                            <button onClick={() => setPdaError(null)} className="ml-auto text-rose-400 hover:text-rose-600">
                              <X className="w-4 h-4" />
                            </button>
                          </motion.div>
                        )}
                      </AnimatePresence>
                      <div className="flex items-center gap-4">
                        <div className="flex-1 relative">
                          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
                          <input 
                            type="text"
                            placeholder="輸入資產編號..."
                            className="w-full pl-10 pr-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all"
                            value={searchCode}
                            onChange={(e) => setSearchCode(e.target.value)}
                            onKeyDown={(e) => e.key === 'Enter' && handleSearch(searchCode)}
                          />
                        </div>
                        <Button onClick={() => handleSearch(searchCode)} className="h-12 px-6">
                          查詢
                        </Button>
                      </div>

                      <div className="relative">
                        <div className="absolute inset-0 flex items-center">
                          <div className="w-full border-t border-slate-200"></div>
                        </div>
                        <div className="relative flex justify-center text-sm">
                          <span className="px-4 bg-white text-slate-500 uppercase tracking-wider font-bold">或</span>
                        </div>
                      </div>

                      <Button 
                        variant="secondary" 
                        className="w-full py-4 text-lg"
                        onClick={() => setScanning(true)}
                      >
                        <Camera className="w-6 h-6" />
                        啟動相機掃描
                      </Button>
                    </Card>

                    {scanning && (
                      <Card className="p-4 relative">
                        <button 
                          onClick={() => setScanning(false)}
                          className="absolute top-2 right-2 z-10 p-2 bg-white rounded-full shadow-lg text-slate-600 hover:text-red-500"
                        >
                          <X className="w-6 h-6" />
                        </button>
                        <div id="reader" className="overflow-hidden rounded-lg bg-slate-900 aspect-square flex items-center justify-center">
                          <div className="text-white/50 text-sm animate-pulse">相機啟動中...</div>
                        </div>
                        <div className="mt-4 space-y-2 text-center">
                          <p className="text-sm text-slate-500">請將 QR Code 置於方框中心進行掃描</p>
                          {window.self !== window.top && (
                            <div className="p-3 bg-rose-50 border border-rose-100 rounded-lg text-xs text-rose-600 font-bold flex items-start gap-2 text-left">
                              <AlertCircle className="w-4 h-4 flex-shrink-0 mt-0.5" />
                              <p>偵測到內嵌視窗限制：手機瀏覽器通常會封鎖內嵌網頁的相機權限。請點擊右上角「在新分頁開啟」圖示後再使用掃描功能。</p>
                            </div>
                          )}
                        </div>
                      </Card>
                    )}

                    {/* Batch List Section */}
                    {batchAssets.length > 0 && (
                      <Card className="p-6 space-y-4">
                        <div className="flex justify-between items-center">
                          <h4 className="font-bold text-slate-900 flex items-center gap-2">
                            <ClipboardList className="w-5 h-5 text-blue-600" />
                            已掃描清單 ({batchAssets.length})
                          </h4>
                          <Button 
                            variant="primary" 
                            className="px-4 py-2 text-sm"
                            onClick={handleSubmitBatch}
                          >
                            提交整批盤點
                          </Button>
                        </div>
                        <div className="space-y-2 max-h-60 overflow-y-auto pr-2">
                          {batchAssets.map((asset, index) => (
                            <div 
                              key={index}
                              onClick={() => {
                                setEditingIndex(index);
                                setPdaRemarks(asset.checkRemark || '');
                                setPdaStatus(asset.checkResult || 'normal');
                                setPdaCustodian(asset.updatedCustodian || '');
                                setPdaLocation(asset.updatedLocation || '');
                                setPdaOffice(asset.updatedOffice || '');
                              }}
                              className={cn(
                                "p-3 rounded-lg border transition-all cursor-pointer flex justify-between items-center",
                                editingIndex === index 
                                  ? "border-blue-500 bg-blue-50" 
                                  : "border-slate-100 hover:border-slate-200 bg-white"
                              )}
                            >
                              <div className="space-y-1">
                                <p className="text-sm font-bold text-slate-900">{asset.assetCode}</p>
                                <p className="text-sm text-slate-600 font-medium">{asset.assetDescription}</p>
                                <div className="flex flex-wrap gap-x-3 gap-y-1 text-sm text-slate-500 font-medium">
                                  <span>保管人: {asset.originalCustodian}</span>
                                  <span>室: {asset.originalOffice}</span>
                                  <span>地點: {asset.originalLocation}</span>
                                </div>
                              </div>
                              <Badge variant={asset.status === 'checked' ? 'success' : 'neutral'}>
                                {asset.status === 'checked' ? '已維護' : '待確認'}
                              </Badge>
                            </div>
                          ))}
                        </div>
                      </Card>
                    )}

                    {/* Edit Form Section */}
                    <AnimatePresence>
                      {editingIndex !== null && (
                        <motion.div
                          initial={{ opacity: 0, y: 20 }}
                          animate={{ opacity: 1, y: 0 }}
                          exit={{ opacity: 0, y: 20 }}
                        >
                          <Card>
                            <div className="bg-blue-600 p-6 text-white">
                              <div className="flex justify-between items-start">
                                <div>
                                  <p className="text-blue-100 text-xs font-bold uppercase tracking-widest mb-1">正在維護資產</p>
                                  <h3 className="text-2xl font-black">{batchAssets[editingIndex].assetCode}</h3>
                                </div>
                                <button onClick={() => setEditingIndex(null)} className="p-2 hover:bg-white/10 rounded-lg">
                                  <X className="w-6 h-6" />
                                </button>
                              </div>
                            </div>
                            <div className="p-6 space-y-6">
                              <div className="grid grid-cols-2 gap-4">
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">資產說明</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].assetDescription}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">科目名稱</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].accountName}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">類別名稱</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].categoryName}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">取得日期</p>
                                  <p className="font-bold text-slate-900">{formatExcelDate(batchAssets[editingIndex].acquisitionDate)}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">取得成本</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].acquisitionCost.toLocaleString()}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">帳面價值</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].bookValue.toLocaleString()}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">數量/單位</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].quantity} {batchAssets[editingIndex].unit}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">成本中心</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].costCenter}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">原保管人</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].originalCustodian}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">原地點</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].originalLocation}</p>
                                </div>
                                <div className="space-y-1">
                                  <p className="text-xs text-slate-400 font-bold uppercase">原室</p>
                                  <p className="font-bold text-slate-900">{batchAssets[editingIndex].originalOffice || '-'}</p>
                                </div>
                              </div>

                              <div className="pt-6 border-t border-slate-100 space-y-4">
                                <div className="space-y-2">
                                  <label className="text-xs text-slate-400 font-bold uppercase">盤點狀態</label>
                                  <div className="grid grid-cols-2 gap-2">
                                    {(['normal', 'abnormal'] as const).map((s) => (
                                      <button
                                        key={s}
                                        onClick={() => setPdaStatus(s)}
                                        className={cn(
                                          "py-2 rounded-lg text-sm font-bold border-2 transition-all",
                                          pdaStatus === s 
                                            ? s === 'normal' ? "bg-emerald-50 border-emerald-500 text-emerald-700" : "bg-amber-50 border-amber-500 text-amber-700"
                                            : "bg-white border-slate-100 text-slate-400 hover:border-slate-200"
                                        )}
                                      >
                                        {s === 'normal' ? '正常' : '異常'}
                                      </button>
                                    ))}
                                  </div>
                                </div>

                                <Input 
                                  label="更新保管人 (若有變動)" 
                                  value={pdaCustodian} 
                                  onChange={(e) => {
                                    setPdaCustodian(e.target.value);
                                    if (e.target.value !== batchAssets[editingIndex!].originalCustodian) setPdaStatus('abnormal');
                                  }}
                                />
                                <Input 
                                  label="更新存放地點 (若有變動)" 
                                  value={pdaLocation} 
                                  onChange={(e) => {
                                    setPdaLocation(e.target.value);
                                    if (e.target.value !== batchAssets[editingIndex!].originalLocation) setPdaStatus('abnormal');
                                  }}
                                />
                                <Input 
                                  label="更新室" 
                                  value={pdaOffice} 
                                  onChange={(e) => {
                                    setPdaOffice(e.target.value);
                                    if (e.target.value !== (batchAssets[editingIndex!].originalOffice || '')) setPdaStatus('abnormal');
                                  }}
                                />
                                <div className="space-y-1">
                                  <label className="block text-sm font-medium text-slate-700">盤點備註</label>
                                  <textarea 
                                    className="w-full px-3 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none min-h-[80px]"
                                    placeholder="輸入異常說明或備註..."
                                    value={pdaRemarks}
                                    onChange={(e) => {
                                      setPdaRemarks(e.target.value);
                                      if (e.target.value.trim() !== '') setPdaStatus('abnormal');
                                    }}
                                  />
                                </div>
                              </div>

                              <Button onClick={handleSaveCurrentEdit} className="w-full py-4 text-lg shadow-xl shadow-blue-500/20">
                                確認此筆明細
                              </Button>
                            </div>
                          </Card>
                        </motion.div>
                      )}
                    </AnimatePresence>
                  </div>
                )}
              </motion.div>
            )}
            {/* Tab: Reports */}
            {activeTab === 'reports' && (
              <motion.div 
                key="reports"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
                className="space-y-6 max-w-7xl mx-auto"
              >
                <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
                  <Card className="p-6 flex flex-col items-center justify-center space-y-4">
                    <h4 className="text-sm font-bold text-slate-400 uppercase tracking-widest">盤點進度</h4>
                    <div className="w-full h-48">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={chartData}
                            innerRadius={60}
                            outerRadius={80}
                            paddingAngle={5}
                            dataKey="value"
                          >
                            {chartData.map((entry, index) => (
                              <Cell key={`cell-${index}`} fill={entry.color} />
                            ))}
                          </Pie>
                          <RechartsTooltip />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                    <div className="text-center">
                      <p className="text-3xl font-black text-slate-900">{Math.round((checkedCount / (assets.length || 1)) * 100)}%</p>
                      <p className="text-xs text-slate-500">已完成盤點</p>
                    </div>
                  </Card>

                  <div className="lg:col-span-3 grid grid-cols-2 md:grid-cols-4 gap-6">
                    {[
                      { label: '總資產數', value: assets.length, icon: Tag, color: 'blue' },
                      { label: '已盤點', value: checkedCount, icon: CheckCircle2, color: 'emerald' },
                      { label: '待盤點', value: pendingCount, icon: History, color: 'slate' },
                      { label: '異常資產', value: abnormalCount, icon: AlertCircle, color: 'rose' }
                    ].map((stat, i) => (
                      <Card key={i} className="p-6 space-y-4">
                        <div className={cn(
                          "w-10 h-10 rounded-lg flex items-center justify-center",
                          stat.color === 'blue' ? "bg-blue-50 text-blue-600" : stat.color === 'emerald' ? "bg-emerald-50 text-emerald-600" : stat.color === 'rose' ? "bg-rose-50 text-rose-600" : "bg-slate-50 text-slate-600"
                        )}>
                          <stat.icon className="w-5 h-5" />
                        </div>
                        <div>
                          <p className="text-sm font-bold text-slate-400 uppercase">{stat.label}</p>
                          <p className="text-2xl font-black text-slate-900">{stat.value}</p>
                        </div>
                      </Card>
                    ))}

                    <Card className="col-span-full p-6">
                      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-6">
                        <div className="space-y-1">
                          <h4 className="font-bold text-slate-900">資產明細清單</h4>
                          <p className="text-sm text-slate-500">即時查看盤點結果與異常備註</p>
                        </div>
                        <div className="flex items-center gap-3 w-full md:w-auto">
                          <div className="relative flex-1 md:w-64">
                            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                            <input 
                              type="text" 
                              placeholder="搜尋編號或名稱..."
                              className="w-full pl-9 pr-4 py-2 rounded-lg border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500/20 outline-none"
                            />
                          </div>
                          <Button variant="outline" className="px-3 py-1.5 text-sm" onClick={handleExportExcel}>
                            <Download className="w-4 h-4" />
                            匯出報表
                          </Button>
                        </div>
                      </div>

                      <div className="overflow-x-auto">
                        <table className="w-full text-sm text-left">
                          <thead className="text-xs text-slate-400 uppercase bg-slate-50">
                            <tr>
                              <th className="px-4 py-3 font-bold">資產編號</th>
                              <th className="px-4 py-3 font-bold">資產說明</th>
                              <th className="px-4 py-3 font-bold">保管人</th>
                              <th className="px-4 py-3 font-bold">地點</th>
                              <th className="px-4 py-3 font-bold">室</th>
                              <th className="px-4 py-3 font-bold">狀態</th>
                              <th className="px-4 py-3 font-bold">結果</th>
                              <th className="px-4 py-3 font-bold">備註</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {assets.map(asset => (
                              <tr key={asset.id} className="hover:bg-slate-50/50 transition-colors">
                                <td className="px-4 py-3 font-mono font-bold text-slate-900">{asset.assetCode}</td>
                                <td className="px-4 py-3 max-w-[200px] truncate">{asset.assetDescription}</td>
                                <td className="px-4 py-3">
                                  <div className="flex flex-col">
                                    <span className="text-slate-900">{asset.updatedCustodian || asset.originalCustodian}</span>
                                    {asset.updatedCustodian && asset.updatedCustodian !== asset.originalCustodian && (
                                      <span className="text-[10px] text-amber-600 font-bold">原: {asset.originalCustodian}</span>
                                    )}
                                  </div>
                                </td>
                                <td className="px-4 py-3">
                                  <div className="flex flex-col">
                                    <span className="text-slate-900">{asset.updatedLocation || asset.originalLocation}</span>
                                    {asset.updatedLocation && asset.updatedLocation !== asset.originalLocation && (
                                      <span className="text-[10px] text-amber-600 font-bold">原: {asset.originalLocation}</span>
                                    )}
                                  </div>
                                </td>
                                <td className="px-4 py-3">
                                  <div className="flex flex-col">
                                    <span className="text-slate-900">{asset.updatedOffice || asset.originalOffice}</span>
                                    {asset.updatedOffice && asset.updatedOffice !== asset.originalOffice && (
                                      <span className="text-[10px] text-amber-600 font-bold">原: {asset.originalOffice}</span>
                                    )}
                                  </div>
                                </td>
                                <td className="px-4 py-3">
                                  <Badge variant={asset.status === 'checked' ? 'success' : 'neutral'}>
                                    {asset.status === 'checked' ? '已盤' : '未盤'}
                                  </Badge>
                                </td>
                                <td className="px-4 py-3">
                                  {asset.checkResult && (
                                    <Badge variant={asset.checkResult === 'normal' ? 'success' : 'warning'}>
                                      {asset.checkResult === 'normal' ? '正常' : '異常'}
                                    </Badge>
                                  )}
                                </td>
                                <td className="px-4 py-3 text-slate-500 max-w-[200px] truncate">{asset.checkRemark || '-'}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </Card>
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </main>

      {/* Toast Notification */}
      <AnimatePresence>
        {toast && (
          <motion.div
            initial={{ opacity: 0, y: 50, scale: 0.9 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, scale: 0.9, y: 20 }}
            className="fixed bottom-6 left-1/2 -translate-x-1/2 z-[100] flex items-center shadow-lg"
          >
            <div className={cn(
              "px-6 py-3 rounded-full flex items-center gap-2 font-medium text-white",
              toast.type === 'success' ? "bg-emerald-600" : 
              toast.type === 'error' ? "bg-rose-600" : "bg-slate-800"
            )}>
              {toast.type === 'success' && <CheckCircle2 className="w-5 h-5" />}
              {toast.type === 'error' && <AlertCircle className="w-5 h-5" />}
              {toast.message}
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* New Plan Modal */}
      <AnimatePresence>
        {showClearConfirm && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="w-full max-w-sm"
            >
              <Card className="border-rose-100">
                <div className="p-6 text-center space-y-4">
                  <div className="w-16 h-16 bg-rose-50 text-rose-600 rounded-full flex items-center justify-center mx-auto">
                    <AlertCircle className="w-8 h-8" />
                  </div>
                  <div className="space-y-2">
                    <h3 className="text-xl font-bold text-slate-900">確定要清除所有資料？</h3>
                    <p className="text-sm text-slate-500 leading-relaxed">
                      這將會永久刪除所有的盤點計畫以及資產明細資料，此動作無法復原。
                    </p>
                  </div>
                  <div className="flex gap-3 pt-2">
                    <Button variant="outline" className="flex-1" onClick={() => setShowClearConfirm(false)} disabled={isClearing}>取消</Button>
                    <Button 
                      className="flex-1 bg-rose-600 hover:bg-rose-700 text-white border-none flex items-center justify-center gap-2" 
                      onClick={handleClearAllData}
                      disabled={isClearing}
                    >
                      {isClearing ? (
                        <>
                          <RefreshCw className="w-4 h-4 animate-spin" />
                          清除中...
                        </>
                      ) : '確定清除'}
                    </Button>
                  </div>
                </div>
              </Card>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {showNewPlanModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="w-full max-w-md"
            >
              <Card>
                <div className="p-6 border-b border-slate-100 flex justify-between items-center">
                  <h3 className="text-xl font-bold text-slate-900">新增盤點計畫</h3>
                  <button onClick={() => setShowNewPlanModal(false)} className="text-slate-400 hover:text-slate-600">
                    <X className="w-6 h-6" />
                  </button>
                </div>
                <div className="p-6 space-y-4">
                  <Input 
                    label="計畫名稱" 
                    placeholder="例如：115年年中資產盤點" 
                    value={newPlan.name}
                    onChange={(e) => setNewPlan({...newPlan, name: e.target.value})}
                  />
                  <Input 
                    label="盤點範圍" 
                    placeholder="例如：全公司、台北總部、資訊處" 
                    value={newPlan.scope}
                    onChange={(e) => setNewPlan({...newPlan, scope: e.target.value})}
                  />
                  <div className="space-y-1">
                    <label className="block text-sm font-medium text-slate-700">計畫說明</label>
                    <textarea 
                      className="w-full px-3 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none min-h-[100px]"
                      placeholder="輸入計畫詳細說明..."
                      value={newPlan.description}
                      onChange={(e) => setNewPlan({...newPlan, description: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="block text-sm font-medium text-slate-700">匯入資產資料 (選填)</label>
                    <div className="relative">
                      <input 
                        type="file" 
                        accept=".xlsx, .xls" 
                        onChange={(e) => setImportFile(e.target.files?.[0] || null)}
                        className="w-full px-3 py-2 rounded-lg border border-slate-200 focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none"
                      />
                    </div>
                    <p className="text-xs text-slate-400">建立計畫時可同步匯入 SAP 匯出的 Excel 檔案</p>
                  </div>
                  <div className="pt-4 flex gap-3">
                    <Button variant="outline" className="flex-1" onClick={() => setShowNewPlanModal(false)} disabled={isCreatingPlan}>取消</Button>
                    <Button className="flex-1" onClick={handleCreatePlan} disabled={isCreatingPlan}>
                      {isCreatingPlan ? (
                        <>
                          <RefreshCw className="w-4 h-4 animate-spin" />
                          建立中...
                        </>
                      ) : '建立計畫'}
                    </Button>
                  </div>
                </div>
              </Card>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
