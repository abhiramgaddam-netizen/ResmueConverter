import { ResumeData } from '@/types/resume';

export interface StoredCandidate {
  id: string;
  name: string;
  size: number;
  status: 'idle' | 'extracting' | 'converting' | 'completed' | 'error';
  extractedText?: string;
  candidateName?: string;
  headerType?: 'role' | 'name';
  docxBase64?: string;
  error?: string;
  data?: ResumeData;
  fileBlob?: Blob;
  updatedAt: number;
}

export interface StoredTemplate {
  key: string; // 'target_template'
  fileName?: string;
  fileBlob?: Blob;
  base64?: string | null;
  text: string;
  isUsingDefault: boolean;
  updatedAt: number;
}

const DB_NAME = 'TecnicsResumeStudioDB';
const DB_VERSION = 1;
const STORE_CANDIDATES = 'candidates';
const STORE_SETTINGS = 'settings';

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof window === 'undefined' || !window.indexedDB) {
      reject(new Error('IndexedDB not supported in this environment'));
      return;
    }

    const request = window.indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      if (!db.objectStoreNames.contains(STORE_CANDIDATES)) {
        db.createObjectStore(STORE_CANDIDATES, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORE_SETTINGS)) {
        db.createObjectStore(STORE_SETTINGS, { keyPath: 'key' });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

/**
 * Save or update all candidates in IndexedDB
 */
export async function dbSaveCandidate(candidate: StoredCandidate): Promise<void> {
  try {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_CANDIDATES, 'readwrite');
      const store = tx.objectStore(STORE_CANDIDATES);
      store.put({ ...candidate, updatedAt: Date.now() });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('Failed to save candidate to IndexedDB:', err);
  }
}

/**
 * Save batch of candidates
 */
export async function dbSaveAllCandidates(candidates: StoredCandidate[]): Promise<void> {
  try {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_CANDIDATES, 'readwrite');
      const store = tx.objectStore(STORE_CANDIDATES);
      candidates.forEach((cand) => {
        store.put({ ...cand, updatedAt: cand.updatedAt || Date.now() });
      });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('Failed to batch save candidates to IndexedDB:', err);
  }
}

/**
 * Load all candidates sorted by insertion order
 */
export async function dbLoadCandidates(): Promise<StoredCandidate[]> {
  try {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_CANDIDATES, 'readonly');
      const store = tx.objectStore(STORE_CANDIDATES);
      const req = store.getAll();
      req.onsuccess = () => {
        const results: StoredCandidate[] = req.result || [];
        results.sort((a, b) => (a.updatedAt || 0) - (b.updatedAt || 0));
        resolve(results);
      };
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    console.warn('Failed to load candidates from IndexedDB:', err);
    return [];
  }
}

/**
 * Delete a candidate by ID
 */
export async function dbDeleteCandidate(id: string): Promise<void> {
  try {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_CANDIDATES, 'readwrite');
      const store = tx.objectStore(STORE_CANDIDATES);
      store.delete(id);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('Failed to delete candidate from IndexedDB:', err);
  }
}

/**
 * Clear all candidates from storage
 */
export async function dbClearCandidates(): Promise<void> {
  try {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_CANDIDATES, 'readwrite');
      const store = tx.objectStore(STORE_CANDIDATES);
      store.clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('Failed to clear candidates from IndexedDB:', err);
  }
}

/**
 * Save target template state
 */
export async function dbSaveTemplate(template: Omit<StoredTemplate, 'key' | 'updatedAt'>): Promise<void> {
  try {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_SETTINGS, 'readwrite');
      const store = tx.objectStore(STORE_SETTINGS);
      store.put({
        key: 'target_template',
        ...template,
        updatedAt: Date.now(),
      });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('Failed to save template to IndexedDB:', err);
  }
}

/**
 * Load saved target template
 */
export async function dbLoadTemplate(): Promise<StoredTemplate | null> {
  try {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_SETTINGS, 'readonly');
      const store = tx.objectStore(STORE_SETTINGS);
      const req = store.get('target_template');
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    console.warn('Failed to load template from IndexedDB:', err);
    return null;
  }
}

/**
 * Reset template to default in storage
 */
export async function dbClearTemplate(): Promise<void> {
  try {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_SETTINGS, 'readwrite');
      const store = tx.objectStore(STORE_SETTINGS);
      store.delete('target_template');
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('Failed to clear template from IndexedDB:', err);
  }
}

/**
 * Complete clear of all studio cached data (Resumes + Template)
 */
export async function dbClearAllStudioData(): Promise<void> {
  try {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      const tx = db.transaction([STORE_CANDIDATES, STORE_SETTINGS], 'readwrite');
      tx.objectStore(STORE_CANDIDATES).clear();
      tx.objectStore(STORE_SETTINGS).clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('Failed to clear all studio data from IndexedDB:', err);
  }
}
