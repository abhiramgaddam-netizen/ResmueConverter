'use client';

import React, { useState, useRef, useEffect } from 'react';
import {
  Upload,
  FileText,
  CheckCircle2,
  AlertCircle,
  Download,
  Settings,
  RefreshCw,
  Trash2,
  Eye,
  FileCheck2,
  Layers,
  Sparkles,
  Building2,
  RotateCcw,
  ShieldCheck,
  Database,
} from 'lucide-react';
import JSZip from 'jszip';
import ResumePreviewModal from '@/components/ResumePreviewModal';
import ConfirmModal from '@/components/ConfirmModal';
import { ResumeData } from '@/types/resume';
import {
  dbSaveCandidate,
  dbSaveAllCandidates,
  dbLoadCandidates,
  dbDeleteCandidate,
  dbClearCandidates,
  dbSaveTemplate,
  dbLoadTemplate,
  dbClearTemplate,
  StoredCandidate,
} from '@/lib/storage';

interface CandidateFileItem {
  id: string;
  file?: File;
  fileBlob?: Blob;
  name: string;
  size: number;
  status: 'idle' | 'extracting' | 'converting' | 'completed' | 'error';
  extractedText?: string;
  candidateName?: string;
  headerType?: 'role' | 'name';
  docxBase64?: string;
  data?: ResumeData;
  error?: string;
  updatedAt?: number;
}

const TECNICS_STANDARD_SPEC = `
TECNICS INTEGRATION TECHNOLOGIES
EXECUTIVE CONSULTANT PROFILE

NAME: [Candidate Full Name]
ROLE: [Target Role / Senior Consultant]
LOCATION: [City, State/Country] | CONTACT: [Email Address] | [Phone Number] | [LinkedIn Profile]

PROFESSIONAL SUMMARY
Results-driven technology consultant with extensive experience delivering mission-critical enterprise solutions. Proven track record in modernizing legacy architectures, leading high-velocity engineering deliverables, and aligning digital transformation initiatives with strategic client goals.

CORE COMPETENCIES & TECHNICAL PROFICIENCIES
- Architecture & Frameworks: Enterprise Cloud Architecture, Microservices, API Engineering
- Technologies & Languages: TypeScript, JavaScript, Python, Java, React, Node.js, Next.js, SQL
- Cloud & Infrastructure: AWS, Azure, GCP, Docker, Kubernetes, CI/CD Automation
- Enterprise Integration: REST APIs, GraphQL, Kafka, Enterprise Service Bus, Database Systems

PROFESSIONAL EXPERIENCE
Client Enterprise Engagement • Location
Senior Consultant / Technical Lead (Start Date - End Date / Present)
- Spearheaded the design and delivery of scalable enterprise services, reducing system latency by 40%.
- Collaborated with executive client stakeholders to define architecture requirements and manage deployments.
- Implemented robust automated verification pipelines, boosting release velocity by 3x.

Prior Client Project / Technology Organization • Location
Software Consultant / Engineer (Start Date - End Date)
- Developed and maintained high-throughput client web applications serving large enterprise user bases.
- Optimized query patterns and distributed data caching, lowering infrastructure operational overhead by 25%.

EDUCATION & CREDENTIALS
- Bachelor of Technology / Bachelor of Science in Computer Science & Engineering

PROFESSIONAL CERTIFICATIONS
- Enterprise Cloud Solutions Architect
`.trim();

export default function TecnicsResumeStudio() {
  // Target Template state
  const [targetFile, setTargetFile] = useState<File | null>(null);
  const [targetTemplateBase64, setTargetTemplateBase64] = useState<string | null>(null);
  const [targetText, setTargetText] = useState<string>(TECNICS_STANDARD_SPEC);
  const [isUsingDefaultTarget, setIsUsingDefaultTarget] = useState<boolean>(true);
  const [isExtractingTarget, setIsExtractingTarget] = useState<boolean>(false);
  const [targetError, setTargetError] = useState<string | null>(null);
  const [showTargetPreview, setShowTargetPreview] = useState<boolean>(false);
  const [templateTab, setTemplateTab] = useState<'spec' | 'placeholders'>('placeholders');

  // Candidate Resumes state
  const [candidates, setCandidates] = useState<CandidateFileItem[]>([]);
  const [isBatchRunning, setIsBatchRunning] = useState<boolean>(false);
  const [currentProcessingIndex, setCurrentProcessingIndex] = useState<number>(-1);
  const [isStorageHydrated, setIsStorageHydrated] = useState<boolean>(false);

  // Preview Modal state
  const [previewCandidate, setPreviewCandidate] = useState<CandidateFileItem | null>(null);
  const [showClearConfirm, setShowClearConfirm] = useState<boolean>(false);

  // Candidate Privacy & Anonymization state
  const [removePersonalInfo, setRemovePersonalInfo] = useState<boolean>(true);

  // Engine Configuration Key state
  const [apiKey, setApiKey] = useState<string>('');
  const [isKeySaved, setIsKeySaved] = useState<boolean>(false);
  const [showKeyConfig, setShowKeyConfig] = useState<boolean>(false);
  const [keyValidationStatus, setKeyValidationStatus] = useState<{
    valid?: boolean;
    error?: string;
  }>({});
  const [isValidatingKey, setIsValidatingKey] = useState<boolean>(false);

  const targetInputRef = useRef<HTMLInputElement>(null);
  const candidateInputRef = useRef<HTMLInputElement>(null);

  // 1. Initial Hydration: Load Key, Saved Template, and Candidates from IndexedDB
  useEffect(() => {
    async function hydrateFromBrowserStorage() {
      if (typeof window === 'undefined') return;

      // Hydrate Privacy setting
      const savedPrivacy = localStorage.getItem('TECNICS_REMOVE_PERSONAL_INFO');
      if (savedPrivacy !== null) {
        setRemovePersonalInfo(savedPrivacy === 'true');
      }

      // Hydrate Key
      const savedKey = localStorage.getItem('TECNICS_AI_KEY') || localStorage.getItem('GEMINI_API_KEY');
      if (savedKey) {
        setApiKey(savedKey);
        setIsKeySaved(true);
        validateKey(savedKey);
      } else {
        validateKey('');
      }

      // Hydrate Target Template from IndexedDB
      try {
        const savedTemplate = await dbLoadTemplate();
        if (savedTemplate) {
          setTargetText(savedTemplate.text || TECNICS_STANDARD_SPEC);
          setTargetTemplateBase64(savedTemplate.base64 || null);
          setIsUsingDefaultTarget(savedTemplate.isUsingDefault);
          if (savedTemplate.fileName) {
            setTargetFile(
              new File(
                savedTemplate.fileBlob ? [savedTemplate.fileBlob] : [''],
                savedTemplate.fileName,
                { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
              )
            );
          }
        }
      } catch (err) {
        console.warn('Failed to restore template from storage:', err);
      }

      // Hydrate Candidates from IndexedDB
      try {
        const storedCandidates = await dbLoadCandidates();
        if (storedCandidates && storedCandidates.length > 0) {
          const restored: CandidateFileItem[] = storedCandidates.map((item) => ({
            id: item.id,
            name: item.name,
            size: item.size,
            status: item.status,
            extractedText: item.extractedText,
            candidateName: item.candidateName,
            headerType: item.headerType,
            docxBase64: item.docxBase64,
            data: item.data,
            error: item.error,
            fileBlob: item.fileBlob,
            file: item.fileBlob
              ? new File([item.fileBlob], item.name, { type: item.fileBlob.type })
              : undefined,
            updatedAt: item.updatedAt,
          }));
          setCandidates(restored);
        }
      } catch (err) {
        console.warn('Failed to restore candidates from storage:', err);
      } finally {
        setIsStorageHydrated(true);
      }
    }

    hydrateFromBrowserStorage();
  }, []);

  const validateKey = async (keyToValidate: string) => {
    setIsValidatingKey(true);
    try {
      const res = await fetch('/api/validate-key', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ apiKey: keyToValidate.trim() || undefined }),
      });
      const data = await res.json();
      if (res.ok && data.valid) {
        setKeyValidationStatus({ valid: true });
      } else {
        setKeyValidationStatus({ valid: false, error: data.error || 'Authorization Required' });
      }
    } catch (err: any) {
      setKeyValidationStatus({ valid: false, error: err.message });
    } finally {
      setIsValidatingKey(false);
    }
  };

  const handleSaveApiKey = () => {
    if (typeof window !== 'undefined') {
      if (apiKey.trim()) {
        localStorage.setItem('TECNICS_AI_KEY', apiKey.trim());
        localStorage.setItem('GEMINI_API_KEY', apiKey.trim());
        setIsKeySaved(true);
        validateKey(apiKey.trim());
      } else {
        localStorage.removeItem('TECNICS_AI_KEY');
        localStorage.removeItem('GEMINI_API_KEY');
        setIsKeySaved(false);
        setKeyValidationStatus({});
      }
    }
  };

  const handleToggleRemovePersonalInfo = (checked: boolean) => {
    setRemovePersonalInfo(checked);
    if (typeof window !== 'undefined') {
      localStorage.setItem('TECNICS_REMOVE_PERSONAL_INFO', String(checked));
    }
  };

  // Convert File to Base64 helper
  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        const base64 = result.split(',')[1] || result;
        resolve(base64);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  };

  // 1. Target Template Upload & Persistence
  const handleTargetUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setTargetFile(file);
    setIsUsingDefaultTarget(false);
    setIsExtractingTarget(true);
    setTargetError(null);

    try {
      const base64 = await fileToBase64(file);
      setTargetTemplateBase64(base64);

      const formData = new FormData();
      formData.append('file', file);

      const res = await fetch('/api/extract-text', {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to extract text from template.');
      }

      setTargetText(data.text);

      // Persist to IndexedDB
      await dbSaveTemplate({
        fileName: file.name,
        fileBlob: file,
        base64: base64,
        text: data.text,
        isUsingDefault: false,
      });
    } catch (err: any) {
      setTargetError(err.message || 'Error processing target template.');
    } finally {
      setIsExtractingTarget(false);
    }
  };

  const handleResetTargetToDefault = async () => {
    setTargetFile(null);
    setTargetTemplateBase64(null);
    setTargetText(TECNICS_STANDARD_SPEC);
    setIsUsingDefaultTarget(true);
    setTargetError(null);
    if (targetInputRef.current) targetInputRef.current.value = '';
    await dbClearTemplate();
  };

  // 2. Candidate Files Upload & Persistence
  const handleCandidateFiles = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const newCandidates: CandidateFileItem[] = Array.from(files).map((f) => ({
      id: Math.random().toString(36).substring(2, 9) + Date.now().toString(36),
      file: f,
      fileBlob: f,
      name: f.name,
      size: f.size,
      status: 'idle',
      updatedAt: Date.now(),
    }));

    setCandidates((prev) => {
      const updated = [...prev, ...newCandidates];
      return updated;
    });

    if (candidateInputRef.current) candidateInputRef.current.value = '';

    // Persist new items to IndexedDB
    const toStore: StoredCandidate[] = newCandidates.map((c) => ({
      id: c.id,
      name: c.name,
      size: c.size,
      status: c.status,
      fileBlob: c.fileBlob,
      updatedAt: c.updatedAt || Date.now(),
    }));
    await dbSaveAllCandidates(toStore);
  };

  const removeCandidate = async (id: string) => {
    if (isBatchRunning) return;
    setCandidates((prev) => prev.filter((c) => c.id !== id));
    await dbDeleteCandidate(id);
    if (previewCandidate?.id === id) {
      setPreviewCandidate(null);
    }
  };

  const clearAllCandidates = () => {
    if (isBatchRunning) return;
    setShowClearConfirm(true);
  };

  const handleConfirmClearAll = async () => {
    setCandidates([]);
    setPreviewCandidate(null);
    await dbClearCandidates();
    setShowClearConfirm(false);
  };

  // Process a candidate
  const processCandidateIndex = async (index: number) => {
    const item = candidates[index];
    if (!item) return;

    setCurrentProcessingIndex(index);

    // Extraction step
    const updatedCandidateExtracting: CandidateFileItem = {
      ...item,
      status: 'extracting',
      error: undefined,
    };

    setCandidates((prev) =>
      prev.map((c, idx) => (idx === index ? updatedCandidateExtracting : c))
    );

    let extractedCandidateText = item.extractedText;

    if (!extractedCandidateText) {
      try {
        const fileToUse = item.file || (item.fileBlob ? new File([item.fileBlob], item.name) : null);
        if (!fileToUse) {
          throw new Error('File data is missing for extraction.');
        }

        const extractFormData = new FormData();
        extractFormData.append('file', fileToUse);

        const extractRes = await fetch('/api/extract-text', {
          method: 'POST',
          body: extractFormData,
        });

        const extractData = await extractRes.json();
        if (!extractRes.ok) {
          throw new Error(extractData.error || 'Failed to extract text from candidate file.');
        }

        extractedCandidateText = extractData.text;
        item.extractedText = extractedCandidateText;

        setCandidates((prev) =>
          prev.map((c, idx) =>
            idx === index ? { ...c, extractedText: extractedCandidateText } : c
          )
        );
      } catch (err: any) {
        const failedItem: CandidateFileItem = {
          ...item,
          status: 'error',
          error: err.message || 'File extraction failed',
          updatedAt: Date.now(),
        };

        setCandidates((prev) =>
          prev.map((c, idx) => (idx === index ? failedItem : c))
        );

        await dbSaveCandidate({
          id: failedItem.id,
          name: failedItem.name,
          size: failedItem.size,
          status: failedItem.status,
          error: failedItem.error,
          fileBlob: failedItem.fileBlob,
          updatedAt: Date.now(),
        });
        return;
      }
    }

    // AI Reformat & DOCX generation using the active template
    setCandidates((prev) =>
      prev.map((c, idx) => (idx === index ? { ...c, status: 'converting' } : c))
    );

    try {
      const convertRes = await fetch('/api/convert-resume', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          candidateText: extractedCandidateText,
          targetFormatText: targetText,
          templateDocxBase64: targetTemplateBase64 || undefined,
          filename: item.name,
          apiKey: apiKey.trim() || undefined,
          removePersonalInfo,
        }),
      });

      const convertData = await convertRes.json();
      if (!convertRes.ok) {
        throw new Error(convertData.error || 'Reformatting process encountered an error.');
      }

      const headerType: 'role' | 'name' =
        convertData.headerType ||
        (targetText.toLowerCase().includes('{{ role }}') || targetText.toLowerCase().includes('{{role}}')
          ? 'role'
          : 'name');

      const completedItem: CandidateFileItem = {
        ...item,
        status: 'completed',
        candidateName: convertData.candidateName,
        headerType,
        docxBase64: convertData.docxBase64,
        data: convertData.data,
        extractedText: extractedCandidateText,
        error: undefined,
        updatedAt: Date.now(),
      };

      setCandidates((prev) =>
        prev.map((c, idx) => (idx === index ? completedItem : c))
      );

      // Persist completed item with docxBase64 and data to IndexedDB
      await dbSaveCandidate({
        id: completedItem.id,
        name: completedItem.name,
        size: completedItem.size,
        status: completedItem.status,
        candidateName: completedItem.candidateName,
        headerType: completedItem.headerType,
        docxBase64: completedItem.docxBase64,
        data: completedItem.data,
        extractedText: completedItem.extractedText,
        fileBlob: completedItem.fileBlob,
        updatedAt: Date.now(),
      });
    } catch (err: any) {
      const errItem: CandidateFileItem = {
        ...item,
        status: 'error',
        error: err.message || 'Transformation failed',
        updatedAt: Date.now(),
      };

      setCandidates((prev) =>
        prev.map((c, idx) => (idx === index ? errItem : c))
      );

      await dbSaveCandidate({
        id: errItem.id,
        name: errItem.name,
        size: errItem.size,
        status: errItem.status,
        error: errItem.error,
        extractedText: errItem.extractedText,
        fileBlob: errItem.fileBlob,
        updatedAt: Date.now(),
      });
    }
  };

  // Batch process
  const handleStartBatch = async () => {
    if (candidates.length === 0 || isBatchRunning) return;

    setIsBatchRunning(true);

    for (let i = 0; i < candidates.length; i++) {
      const item = candidates[i];
      if (item.status === 'completed') continue;

      await processCandidateIndex(i);
      await new Promise((resolve) => setTimeout(resolve, 800));
    }

    setIsBatchRunning(false);
    setCurrentProcessingIndex(-1);
  };

  const handleRetryCandidate = async (index: number) => {
    if (isBatchRunning) return;
    setIsBatchRunning(true);
    await processCandidateIndex(index);
    setIsBatchRunning(false);
    setCurrentProcessingIndex(-1);
  };

  // Download Single Word Document
  const downloadSingleDocx = (item: CandidateFileItem) => {
    if (!item.docxBase64) return;
    const byteCharacters = atob(item.docxBase64);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });

    const safeName = (item.candidateName || item.name.replace(/\.[^/.]+$/, '')).replace(/[/\\?%*:|"<>]/g, '_');
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Tecnics_${safeName}_Standardized.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Download All as ZIP
  const downloadAllZip = async () => {
    const completedItems = candidates.filter((c) => c.status === 'completed' && c.docxBase64);
    if (completedItems.length === 0) return;

    const zip = new JSZip();

    completedItems.forEach((item, index) => {
      const byteCharacters = atob(item.docxBase64!);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const safeName = (item.candidateName || `Candidate_${index + 1}`).replace(/[/\\?%*:|"<>]/g, '_');
      zip.file(`Tecnics_${safeName}_Standardized.docx`, byteArray);
    });

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Tecnics_Standardized_Resumes_${new Date().toISOString().slice(0, 10)}.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const completedCount = candidates.filter((c) => c.status === 'completed').length;
  const errorCount = candidates.filter((c) => c.status === 'error').length;
  const progressPercent = candidates.length > 0 ? Math.round((completedCount / candidates.length) * 100) : 0;

  return (
    <div className="min-h-screen bg-white text-slate-900 flex flex-col justify-between selection:bg-blue-600 selection:text-white">
      <div>
        {/* Simple & Elegant Header */}
        <header className="border-b border-slate-200 bg-white sticky top-0 z-40">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
            {/* Brand Logo & Title */}
            <div className="flex items-center space-x-3.5">
              <div className="w-10 h-10 rounded-lg bg-slate-900 flex items-center justify-center text-white shadow-sm">
                <Building2 className="w-5 h-5 text-blue-400" />
              </div>
              <div>
                <div className="flex items-center space-x-2">
                  <span className="font-bold text-lg tracking-tight text-slate-900">TECNICS</span>
                  <span className="text-xs font-semibold px-2 py-0.5 rounded bg-slate-100 text-slate-700 border border-slate-200">
                    RESUME STUDIO
                  </span>
                </div>
                <p className="text-xs text-slate-500 font-normal">
                  Tecnics Integration Technologies Private Limited
                </p>
              </div>
            </div>

            {/* Storage Status & Engine Key Settings */}
            <div className="flex items-center space-x-3">
              {/* Persistence indicator */}
              <div
                className="hidden md:flex items-center text-xs bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-full text-slate-600"
                title="Resumes and templates are cached in browser storage and persist on refresh until cleared."
              >
                <Database className="w-3.5 h-3.5 text-blue-600 mr-1.5" />
                <span>Auto-Saved in Browser</span>
              </div>

              <div className="hidden sm:flex items-center text-xs bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-full text-slate-700">
                <span
                  className={`w-2 h-2 rounded-full mr-2 ${
                    keyValidationStatus.valid ? 'bg-emerald-500' : 'bg-slate-400'
                  }`}
                ></span>
                <span className="font-medium">
                  {keyValidationStatus.valid ? 'Engine Online' : 'Engine Ready'}
                </span>
              </div>

              <button
                onClick={() => setShowKeyConfig(!showKeyConfig)}
                className="flex items-center space-x-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 transition-colors shadow-sm"
              >
                <Settings className="w-3.5 h-3.5 text-slate-500" />
                <span>Settings</span>
              </button>
            </div>
          </div>

          {/* Engine Key Configuration Bar */}
          {showKeyConfig && (
            <div className="border-t border-slate-200 bg-slate-50 px-4 py-3 sm:px-8">
              <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-3 text-xs">
                <div className="text-slate-600">
                  <span className="font-medium text-slate-800">Engine Authorization:</span> Provide your authorization key or rely on server defaults.
                </div>
                <div className="flex items-center space-x-2 w-full sm:w-auto">
                  <input
                    type="password"
                    placeholder="Enter Engine Key..."
                    value={apiKey}
                    onChange={(e) => setApiKey(e.target.value)}
                    className="bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-xs text-slate-800 w-full sm:w-72 focus:outline-none focus:ring-1 focus:ring-blue-500 font-mono"
                  />
                  <button
                    onClick={handleSaveApiKey}
                    disabled={isValidatingKey}
                    className="bg-slate-900 hover:bg-slate-800 text-white font-medium px-3.5 py-1.5 rounded-lg whitespace-nowrap transition-colors flex items-center space-x-1.5 text-xs"
                  >
                    {isValidatingKey ? (
                      <>
                        <RefreshCw className="w-3 h-3 animate-spin" />
                        <span>Verifying...</span>
                      </>
                    ) : (
                      <span>Save Key</span>
                    )}
                  </button>
                </div>
              </div>
            </div>
          )}
        </header>

        {/* Main Content Area */}
        <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-8 pb-12">
          {/* Executive Overview Header */}
          <div className="border-b border-slate-200 pb-8 mb-8">
            <div className="max-w-3xl">
              <span className="text-xs font-semibold uppercase tracking-wider text-blue-600 bg-blue-50 px-2.5 py-1 rounded">
                Tecnics Enterprise Portal
              </span>
              <h2 className="text-2xl sm:text-3xl font-bold tracking-tight text-slate-900 mt-2.5 mb-2">
                Candidate Resume Standardization
              </h2>
              <p className="text-slate-600 text-sm sm:text-base leading-relaxed">
                Standardize multiple candidate resumes into our corporate Word template (`.docx`) with verified factual accuracy.
                All processed profiles are previewable directly in the browser and automatically preserved across refreshes.
              </p>
            </div>
          </div>

          {/* Candidate Privacy & Blind Profile Banner */}
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 mb-8 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-xs">
            <div className="flex items-center space-x-3">
              <div className="w-9 h-9 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-700 flex-shrink-0">
                <ShieldCheck className="w-5 h-5 text-emerald-600" />
              </div>
              <div>
                <div className="flex items-center space-x-2">
                  <span className="text-xs font-bold text-slate-900">Remove Candidate Personal Details</span>
                  <span className="text-[10px] font-semibold px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">
                    Client Blind Profile
                  </span>
                </div>
                <p className="text-[11px] text-slate-500 mt-0.5">
                  Excludes candidate email, phone number, address/location, LinkedIn, and GitHub from generated Word documents and previews.
                </p>
              </div>
            </div>

            <label className="flex items-center space-x-2 cursor-pointer select-none flex-shrink-0 bg-white border border-slate-300 hover:border-slate-400 px-3.5 py-2 rounded-lg shadow-xs transition-colors">
              <input
                type="checkbox"
                checked={removePersonalInfo}
                onChange={(e) => handleToggleRemovePersonalInfo(e.target.checked)}
                className="w-4 h-4 text-blue-600 rounded border-slate-300 focus:ring-blue-500 cursor-pointer"
              />
              <span className="text-xs font-semibold text-slate-800">
                Omit Personal Details
              </span>
            </label>
          </div>

          {/* Configuration Grid: 2 Clean Steps */}
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 mb-8">
            {/* STEP 1: Benchmark Template (5 Cols) */}
            <div className="lg:col-span-5 bg-white border border-slate-200 rounded-xl p-6 shadow-sm flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between mb-3">
                  <div className="flex items-center space-x-2">
                    <span className="w-6 h-6 rounded-md bg-slate-100 text-slate-700 text-xs font-bold flex items-center justify-center">
                      1
                    </span>
                    <h3 className="font-semibold text-slate-900 text-sm">Corporate DOCX Template</h3>
                  </div>
                  {isUsingDefaultTarget ? (
                    <span className="text-xs bg-slate-100 text-slate-600 px-2.5 py-0.5 rounded-full font-medium">
                      Tecnics Official Template
                    </span>
                  ) : (
                    <button
                      onClick={handleResetTargetToDefault}
                      className="text-xs text-blue-600 hover:text-blue-800 font-medium"
                    >
                      Reset to Default
                    </button>
                  )}
                </div>

                <p className="text-xs text-slate-500 mb-4 leading-relaxed">
                  Upload your customized Word template (`.docx`) containing your company logo &amp; header, or use the official Tecnics template.
                </p>

                {/* Upload Drop Zone */}
                <input
                  ref={targetInputRef}
                  type="file"
                  accept=".docx,.doc,.pdf,.txt"
                  onChange={handleTargetUpload}
                  className="hidden"
                  id="target-upload"
                />

                <label
                  htmlFor="target-upload"
                  className={`border border-dashed rounded-lg p-6 flex flex-col items-center justify-center text-center cursor-pointer transition-all ${
                    targetFile
                      ? 'border-emerald-300 bg-emerald-50/30 hover:bg-emerald-50/60'
                      : 'border-slate-300 hover:border-blue-400 bg-slate-50/50 hover:bg-slate-50'
                  }`}
                >
                  {isExtractingTarget ? (
                    <div className="flex flex-col items-center py-2">
                      <RefreshCw className="w-6 h-6 text-blue-600 animate-spin mb-1.5" />
                      <span className="text-xs font-medium text-slate-700">Loading template binary...</span>
                    </div>
                  ) : targetFile ? (
                    <div className="flex flex-col items-center py-1">
                      <FileCheck2 className="w-6 h-6 text-emerald-600 mb-1.5" />
                      <span className="text-xs font-semibold text-slate-900 truncate max-w-xs">{targetFile.name}</span>
                      <span className="text-[11px] text-emerald-700 mt-0.5">Template Loaded &amp; Active ✓</span>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center py-1">
                      <Upload className="w-6 h-6 text-slate-400 mb-1.5" />
                      <span className="text-xs font-medium text-slate-700">Upload Target Template (.docx)</span>
                      <span className="text-[11px] text-slate-400 mt-0.5">Retains exact logo, margins, and header lines</span>
                    </div>
                  )}
                </label>

                {targetError && (
                  <div className="mt-3 p-2.5 bg-red-50 border border-red-200 rounded-lg text-xs text-red-600 flex items-start space-x-2">
                    <AlertCircle className="w-4 h-4 flex-shrink-0 mt-0.5" />
                    <span>{targetError}</span>
                  </div>
                )}
              </div>

              <div className="mt-5 pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
                <span className="text-slate-500">
                  {targetFile ? 'Custom Template Active' : 'Official Tecnics Logo & Format'}
                </span>
                <button
                  onClick={() => setShowTargetPreview(!showTargetPreview)}
                  className="text-blue-600 hover:text-blue-800 font-medium inline-flex items-center"
                >
                  <Eye className="w-3.5 h-3.5 mr-1" />
                  {showTargetPreview ? 'Hide Spec' : 'Inspect Spec'}
                </button>
              </div>
            </div>

            {/* STEP 2: Candidate Resumes (7 Cols) */}
            <div className="lg:col-span-7 bg-white border border-slate-200 rounded-xl p-6 shadow-sm flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between mb-3">
                  <div className="flex items-center space-x-2">
                    <span className="w-6 h-6 rounded-md bg-slate-100 text-slate-700 text-xs font-bold flex items-center justify-center">
                      2
                    </span>
                    <h3 className="font-semibold text-slate-900 text-sm">Candidate Resumes</h3>
                  </div>

                  {candidates.length > 0 && (
                    <div className="flex items-center space-x-2">
                      <span className="text-xs font-medium bg-slate-100 text-slate-700 px-2.5 py-0.5 rounded-full">
                        {candidates.length} {candidates.length === 1 ? 'Resume' : 'Resumes'}
                      </span>
                      {!isBatchRunning && (
                        <button
                          onClick={clearAllCandidates}
                          className="text-xs text-slate-500 hover:text-red-600 transition-colors flex items-center space-x-1 px-2 py-0.5 rounded hover:bg-red-50"
                          title="Clear all resumes and local browser cache"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          <span>Clear All</span>
                        </button>
                      )}
                    </div>
                  )}
                </div>

                <p className="text-xs text-slate-500 mb-4 leading-relaxed">
                  Select multiple candidate resumes at once. Supported formats: PDF and Word DOCX.
                </p>

                {/* Candidate Dropzone */}
                <input
                  ref={candidateInputRef}
                  type="file"
                  multiple
                  accept=".docx,.doc,.pdf,.txt"
                  onChange={handleCandidateFiles}
                  className="hidden"
                  id="candidates-upload"
                  disabled={isBatchRunning}
                />

                <label
                  htmlFor="candidates-upload"
                  className={`border border-dashed rounded-lg p-6 flex flex-col items-center justify-center text-center cursor-pointer transition-all ${
                    isBatchRunning
                      ? 'opacity-50 cursor-not-allowed bg-slate-50 border-slate-200'
                      : 'border-slate-300 hover:border-blue-500 bg-slate-50/50 hover:bg-slate-50'
                  }`}
                >
                  <Layers className="w-7 h-7 text-blue-600 mb-1.5" />
                  <span className="text-xs font-medium text-slate-800">
                    Click or drag &amp; drop multiple resumes here
                  </span>
                  <span className="text-[11px] text-slate-500 mt-0.5">
                    Select multiple files simultaneously (.pdf, .docx)
                  </span>
                </label>
              </div>

              {/* Action Bar */}
              <div className="mt-5 pt-3 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3">
                <div className="text-xs text-slate-500">
                  {candidates.length === 0 ? (
                    <span>No resumes uploaded</span>
                  ) : (
                    <span>
                      Total: <strong className="text-slate-800">{candidates.length}</strong> | Processed:{' '}
                      <strong className="text-emerald-600">{completedCount}</strong>
                    </span>
                  )}
                </div>

                <button
                  onClick={handleStartBatch}
                  disabled={isBatchRunning || candidates.length === 0}
                  className={`w-full sm:w-auto px-6 py-2.5 rounded-lg font-medium text-xs shadow-sm flex items-center justify-center space-x-2 transition-all ${
                    isBatchRunning || candidates.length === 0
                      ? 'bg-slate-100 text-slate-400 cursor-not-allowed border border-slate-200'
                      : 'bg-blue-600 hover:bg-blue-700 text-white active:scale-98'
                  }`}
                >
                  {isBatchRunning ? (
                    <>
                      <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                      <span>Processing Batch...</span>
                    </>
                  ) : (
                    <>
                      <Sparkles className="w-3.5 h-3.5" />
                      <span>Standardize {candidates.length > 0 ? `${candidates.length} Resumes` : 'All'}</span>
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>

          {/* Template Inspection Drawer */}
          {showTargetPreview && (
            <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm mb-8 animate-fadeIn">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4 pb-3 border-b border-slate-100">
                <div className="flex items-center space-x-2">
                  <button
                    onClick={() => setTemplateTab('placeholders')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                      templateTab === 'placeholders'
                        ? 'bg-blue-50 text-blue-700 border border-blue-200'
                        : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                    }`}
                  >
                    Supported Word Placeholders (.docx)
                  </button>
                  <button
                    onClick={() => setTemplateTab('spec')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                      templateTab === 'spec'
                        ? 'bg-blue-50 text-blue-700 border border-blue-200'
                        : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                    }`}
                  >
                    Raw Benchmark Specification
                  </button>
                </div>
                <button
                  onClick={() => setShowTargetPreview(false)}
                  className="text-xs text-slate-500 hover:text-slate-800 self-end sm:self-auto font-medium"
                >
                  Close
                </button>
              </div>

              {templateTab === 'placeholders' ? (
                <div className="space-y-4">
                  <div className="bg-blue-50/50 border border-blue-100 rounded-lg p-3 text-xs text-blue-900 leading-relaxed">
                    💡 <strong>Smart Heading Detection:</strong> If you design your own styled section heading in Microsoft Word above a placeholder (e.g. typing <em>&quot;CAREER HIGHLIGHTS&quot;</em> directly above <code className="bg-blue-100/70 px-1 py-0.5 rounded font-mono text-blue-800">{"{{ summary }}"}</code>), the converter will automatically omit the default heading to prevent repetition!
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                    <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/40">
                      <div className="font-semibold text-slate-900 mb-1.5 flex items-center justify-between">
                        <span>Candidate Header &amp; Role</span>
                        <span className="text-[10px] text-slate-500 font-normal">Header or Body</span>
                      </div>
                      <ul className="space-y-1.5 text-slate-700">
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ Role }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Job title / specialization</span>
                        </li>
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ candidate_name }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Full candidate name</span>
                        </li>
                      </ul>
                    </div>

                    <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/40">
                      <div className="font-semibold text-slate-900 mb-1.5 flex items-center justify-between">
                        <span>Modular Section Headings</span>
                        <span className="text-[10px] text-slate-500 font-normal">Individual Blocks</span>
                      </div>
                      <ul className="space-y-1.5 text-slate-700">
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ summary }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Executive Overview</span>
                        </li>
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ skills }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Categorized competencies</span>
                        </li>
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ experience }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Work history &amp; bullets</span>
                        </li>
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ projects }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Key projects &amp; tech</span>
                        </li>
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ education }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Degrees &amp; universities</span>
                        </li>
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ certifications }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Certificates &amp; licenses</span>
                        </li>
                      </ul>
                    </div>

                    <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/40">
                      <div className="font-semibold text-slate-900 mb-1.5 flex items-center justify-between">
                        <span>Contact Information</span>
                        <span className="text-[10px] text-slate-500 font-normal">Omitted in blind mode</span>
                      </div>
                      <ul className="space-y-1.5 text-slate-700">
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ contact_info }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Formatted contact bar</span>
                        </li>
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ email }}"} / {"{{ phone }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Inline fields</span>
                        </li>
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ location }}"} / {"{{ linkedin }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">City/State &amp; profile</span>
                        </li>
                      </ul>
                    </div>

                    <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/40">
                      <div className="font-semibold text-slate-900 mb-1.5 flex items-center justify-between">
                        <span>Full Catch-All Placeholder</span>
                        <span className="text-[10px] text-slate-500 font-normal">Auto Remaining</span>
                      </div>
                      <ul className="space-y-1.5 text-slate-700">
                        <li className="flex items-start justify-between">
                          <code className="font-mono text-blue-600 font-semibold bg-white px-1.5 py-0.5 rounded border border-slate-200">{"{{ content }}"}</code>
                          <span className="text-slate-500 text-[11px] ml-2">Renders all unplaced sections</span>
                        </li>
                      </ul>
                      <p className="text-[11px] text-slate-500 mt-2 leading-relaxed">
                        If you place <code className="bg-slate-200 px-1 rounded font-mono">{"{{ summary }}"}</code> separately and also include <code className="bg-slate-200 px-1 rounded font-mono">{"{{ content }}"}</code>, <code className="bg-slate-200 px-1 rounded font-mono">{"{{ content }}"}</code> only outputs the remaining sections without repeating summary!
                      </p>
                    </div>
                  </div>
                </div>
              ) : (
                <pre className="text-xs bg-slate-50 text-slate-800 p-4 rounded-lg overflow-x-auto max-h-64 font-mono leading-relaxed whitespace-pre-wrap border border-slate-200">
                  {targetText}
                </pre>
              )}
            </div>
          )}

          {/* Batch Progress Bar */}
          {candidates.length > 0 && (
            <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm mb-8">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-3">
                <div>
                  <h3 className="font-semibold text-slate-900 text-sm">Batch Progress</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    {isBatchRunning
                      ? `Converting resume ${currentProcessingIndex + 1} of ${candidates.length}...`
                      : completedCount === candidates.length && candidates.length > 0
                      ? 'Batch complete! All resumes standardized and stored in browser.'
                      : `${completedCount} of ${candidates.length} completed`}
                  </p>
                </div>

                {completedCount > 0 && (
                  <button
                    onClick={downloadAllZip}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-4 py-2 rounded-lg shadow-sm flex items-center justify-center space-x-1.5 transition-colors"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Download All (.ZIP)</span>
                  </button>
                )}
              </div>

              {/* Progress Line */}
              <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden mb-2">
                <div
                  className="bg-blue-600 h-full rounded-full transition-all duration-300 ease-out"
                  style={{ width: `${progressPercent}%` }}
                ></div>
              </div>

              <div className="flex items-center justify-between text-xs text-slate-500">
                <span>{progressPercent}% Complete</span>
                <div className="flex items-center space-x-3">
                  <span className="flex items-center text-emerald-600 font-medium">
                    <CheckCircle2 className="w-3.5 h-3.5 mr-1" /> {completedCount} Done
                  </span>
                  {errorCount > 0 && (
                    <span className="flex items-center text-red-600 font-medium">
                      <AlertCircle className="w-3.5 h-3.5 mr-1" /> {errorCount} Attention
                    </span>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* Resumes Queue Table */}
          {candidates.length > 0 && (
            <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden mb-8">
              <div className="px-5 py-3.5 border-b border-slate-100 bg-slate-50 flex items-center justify-between">
                <div className="flex items-center space-x-2">
                  <h3 className="font-semibold text-slate-900 text-xs">Resumes Queue ({candidates.length})</h3>
                  <span className="text-[11px] text-slate-400 hidden sm:inline">• Persisted until cleared</span>
                </div>
                <span className="text-xs text-slate-500">Corporate Word Output</span>
              </div>

              <div className="divide-y divide-slate-100 max-h-[500px] overflow-y-auto">
                {candidates.map((item, index) => {
                  const isCurrent = isBatchRunning && currentProcessingIndex === index;
                  const canView = item.data || item.status === 'completed' || Boolean(item.extractedText);

                  return (
                    <div
                      key={item.id}
                      className={`px-5 py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 transition-colors ${
                        isCurrent ? 'bg-blue-50/50' : 'hover:bg-slate-50/60'
                      }`}
                    >
                      <div className="flex items-center space-x-3 min-w-0">
                        <div className="w-7 h-7 rounded-md bg-slate-100 flex items-center justify-center flex-shrink-0 text-slate-600">
                          <FileText className="w-3.5 h-3.5" />
                        </div>
                        <div className="min-w-0">
                          <div className="flex items-center space-x-1.5">
                            <p className="text-xs font-semibold text-slate-900 truncate">
                              {item.candidateName || item.name}
                            </p>
                            {item.candidateName && (
                              <span className="text-[11px] text-slate-400 truncate">({item.name})</span>
                            )}
                          </div>
                          <p className="text-[11px] text-slate-400">{(item.size / 1024).toFixed(1)} KB</p>
                          {item.error && (
                            <p className="text-xs text-red-600 mt-1 max-w-xl font-mono bg-red-50 p-1.5 rounded border border-red-200">
                              {item.error}
                            </p>
                          )}
                        </div>
                      </div>

                      <div className="flex items-center space-x-2 self-end sm:self-center">
                        {/* Status Badges */}
                        {item.status === 'idle' && (
                          <span className="text-xs bg-slate-100 text-slate-600 px-2.5 py-0.5 rounded-full font-medium">
                            Queued
                          </span>
                        )}
                        {item.status === 'extracting' && (
                          <span className="text-xs bg-amber-50 text-amber-700 border border-amber-200 px-2.5 py-0.5 rounded-full font-medium flex items-center">
                            <RefreshCw className="w-3 h-3 animate-spin mr-1" />
                            Reading
                          </span>
                        )}
                        {item.status === 'converting' && (
                          <span className="text-xs bg-blue-50 text-blue-700 border border-blue-200 px-2.5 py-0.5 rounded-full font-medium flex items-center">
                            <RefreshCw className="w-3 h-3 animate-spin mr-1" />
                            Formatting...
                          </span>
                        )}
                        {item.status === 'completed' && (
                          <span className="text-xs bg-emerald-50 text-emerald-700 border border-emerald-200 px-2.5 py-0.5 rounded-full font-medium flex items-center">
                            <CheckCircle2 className="w-3 h-3 mr-1" />
                            Formatted
                          </span>
                        )}
                        {item.status === 'error' && (
                          <div className="flex items-center space-x-1.5">
                            <span
                              className="text-xs bg-red-50 text-red-700 border border-red-200 px-2 py-0.5 rounded-full font-medium flex items-center"
                              title={item.error}
                            >
                              <AlertCircle className="w-3 h-3 mr-1" />
                              Attention
                            </span>
                            {!isBatchRunning && (
                              <button
                                onClick={() => handleRetryCandidate(index)}
                                className="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-2 py-0.5 rounded flex items-center space-x-1 border border-slate-200"
                                title="Retry conversion"
                              >
                                <RotateCcw className="w-3 h-3" />
                                <span>Retry</span>
                              </button>
                            )}
                          </div>
                        )}

                        {/* In-Browser VIEW Option */}
                        {canView && (
                          <button
                            onClick={() => setPreviewCandidate(item)}
                            className="bg-white hover:bg-blue-50 border border-slate-300 hover:border-blue-400 text-slate-700 hover:text-blue-700 text-xs font-medium px-2.5 py-1 rounded flex items-center space-x-1 transition-colors shadow-sm"
                            title="View formatted profile in browser"
                          >
                            <Eye className="w-3.5 h-3.5 text-blue-600" />
                            <span>View</span>
                          </button>
                        )}

                        {/* Download Single DOCX */}
                        {item.status === 'completed' && item.docxBase64 && (
                          <button
                            onClick={() => downloadSingleDocx(item)}
                            className="bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-800 text-xs font-medium px-2.5 py-1 rounded flex items-center space-x-1 transition-colors"
                            title="Download Word Document (.docx)"
                          >
                            <Download className="w-3 h-3" />
                            <span>.docx</span>
                          </button>
                        )}

                        {!isBatchRunning && (
                          <button
                            onClick={() => removeCandidate(item.id)}
                            className="text-slate-400 hover:text-red-600 p-1 transition-colors"
                            title="Remove from batch and storage"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Compliance & Governance Note */}
          <div className="border border-slate-200 rounded-xl p-5 bg-slate-50/70 text-xs text-slate-600">
            <h4 className="font-semibold text-slate-800 text-xs mb-1.5 flex items-center">
              <ShieldCheck className="w-4 h-4 text-blue-600 mr-1.5" />
              Corporate Data Governance Standards
            </h4>
            <p className="leading-relaxed text-slate-500 mb-2">
              The standardization engine maps candidate career facts directly into the corporate specification. Work
              histories, dates, and qualifications are preserved without modification or artificial generation.
              Resumes and templates are cached securely in your browser session and remain available until you choose to clear them.
            </p>
            <div className="flex flex-wrap gap-4 text-slate-400 text-[11px] pt-2 border-t border-slate-200">
              <span>• In-Browser Instant Profile Viewer</span>
              <span>• Persistent Browser Cache on Refresh</span>
              <span>• Executive Word (.docx) Format</span>
              <span>• Batch Multi-Candidate Ingestion</span>
              <span>• Zero Fact Embellishment</span>
              <span>• Instant ZIP Packaging</span>
            </div>
          </div>
        </main>
      </div>

      {/* In-Browser Resume Preview Modal */}
      {previewCandidate && (
        <ResumePreviewModal
          isOpen={Boolean(previewCandidate)}
          onClose={() => setPreviewCandidate(null)}
          candidateName={previewCandidate.candidateName || previewCandidate.name}
          originalFilename={previewCandidate.name}
          data={previewCandidate.data}
          extractedText={previewCandidate.extractedText}
          docxBase64={previewCandidate.docxBase64}
          removePersonalInfo={removePersonalInfo}
          headerType={
            previewCandidate.headerType ||
            (targetText.toLowerCase().includes('{{ role }}') || targetText.toLowerCase().includes('{{role}}')
              ? 'role'
              : 'name')
          }
          onDownloadDocx={
            previewCandidate.docxBase64
              ? () => downloadSingleDocx(previewCandidate)
              : undefined
          }
        />
      )}

      {/* Corporate Confirmation Modal for Clearing Batch & Cache */}
      <ConfirmModal
        isOpen={showClearConfirm}
        onClose={() => setShowClearConfirm(false)}
        onConfirm={handleConfirmClearAll}
        title="Clear All Resumes & Storage?"
        description={`Are you sure you want to remove all ${candidates.length} candidate resumes? This will permanently remove your active queue and clear the cached resumes from your browser storage.`}
        confirmText="Yes, Clear All"
        cancelText="Cancel"
        danger={true}
      />

      {/* Corporate Footer */}
      <footer className="border-t border-slate-200 bg-white py-6 px-4 sm:px-6 lg:px-8 mt-auto">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500">
          <div className="flex items-center space-x-2">
            <Building2 className="w-4 h-4 text-slate-600" />
            <span className="font-semibold text-slate-700">
              Tecnics Integration Technologies Private Limited
            </span>
          </div>
          <div>
            <span>© {new Date().getFullYear()} Tecnics Integration Technologies Private Limited. All Rights Reserved.</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
