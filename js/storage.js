/**
 * IndexedDB 存储封装
 * 提供模板和历史记录的持久化存储功能
 */

const DB_NAME = 'WordTemplateToolDB';
const DB_VERSION = 1;

// 存储名称
const STORES = {
    TEMPLATES: 'templates',
    RECORDS: 'records'
};

// 数据库实例
let db = null;

/**
 * 初始化数据库
 * @returns {Promise<IDBDatabase>}
 */
function initDB() {
    return new Promise((resolve, reject) => {
        if (db) {
            resolve(db);
            return;
        }

        const request = indexedDB.open(DB_NAME, DB_VERSION);

        request.onerror = () => {
            reject(new Error('无法打开数据库'));
        };

        request.onsuccess = (event) => {
            db = event.target.result;
            resolve(db);
        };

        request.onupgradeneeded = (event) => {
            const database = event.target.result;

            // 创建模板存储
            if (!database.objectStoreNames.contains(STORES.TEMPLATES)) {
                const templateStore = database.createObjectStore(STORES.TEMPLATES, { keyPath: 'id' });
                templateStore.createIndex('name', 'name', { unique: false });
                templateStore.createIndex('createdAt', 'createdAt', { unique: false });
                templateStore.createIndex('updatedAt', 'updatedAt', { unique: false });
            }

            // 创建记录存储
            if (!database.objectStoreNames.contains(STORES.RECORDS)) {
                const recordStore = database.createObjectStore(STORES.RECORDS, { keyPath: 'id' });
                recordStore.createIndex('templateId', 'templateId', { unique: false });
                recordStore.createIndex('createdAt', 'createdAt', { unique: false });
                recordStore.createIndex('updatedAt', 'updatedAt', { unique: false });
            }
        };
    });
}

/**
 * 生成唯一 ID
 * @returns {string}
 */
function generateId() {
    return 'id_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2, 9);
}

/**
 * 获取当前时间戳
 * @returns {string}
 */
function getTimestamp() {
    return new Date().toISOString();
}

// ==========================================
// 模板存储操作
// ==========================================

/**
 * 保存模板
 * @param {Object} template - 模板数据
 * @returns {Promise<Object>}
 */
async function saveTemplate(template) {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.TEMPLATES], 'readwrite');
        const store = transaction.objectStore(STORES.TEMPLATES);
        
        const now = getTimestamp();
        const isNew = !template.id;
        
        const templateData = {
            ...template,
            id: template.id || generateId(),
            createdAt: template.createdAt || now,
            updatedAt: now
        };

        const request = store.put(templateData);

        request.onsuccess = () => {
            resolve(templateData);
        };

        request.onerror = () => {
            reject(new Error('保存模板失败'));
        };
    });
}

/**
 * 获取所有模板
 * @returns {Promise<Array>}
 */
async function getAllTemplates() {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.TEMPLATES], 'readonly');
        const store = transaction.objectStore(STORES.TEMPLATES);
        const request = store.getAll();

        request.onsuccess = () => {
            // 按更新时间倒序排列
            const templates = request.result.sort((a, b) => 
                new Date(b.updatedAt) - new Date(a.updatedAt)
            );
            resolve(templates);
        };

        request.onerror = () => {
            reject(new Error('获取模板列表失败'));
        };
    });
}

/**
 * 根据 ID 获取模板
 * @param {string} id - 模板 ID
 * @returns {Promise<Object|null>}
 */
async function getTemplateById(id) {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.TEMPLATES], 'readonly');
        const store = transaction.objectStore(STORES.TEMPLATES);
        const request = store.get(id);

        request.onsuccess = () => {
            resolve(request.result || null);
        };

        request.onerror = () => {
            reject(new Error('获取模板失败'));
        };
    });
}

/**
 * 删除模板
 * @param {string} id - 模板 ID
 * @returns {Promise<void>}
 */
async function deleteTemplate(id) {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.TEMPLATES], 'readwrite');
        const store = transaction.objectStore(STORES.TEMPLATES);
        const request = store.delete(id);

        request.onsuccess = () => {
            resolve();
        };

        request.onerror = () => {
            reject(new Error('删除模板失败'));
        };
    });
}

/**
 * 检查模板名称是否已存在
 * @param {string} name - 模板名称
 * @param {string} excludeId - 排除的模板 ID（用于编辑时）
 * @returns {Promise<boolean>}
 */
async function isTemplateNameExists(name, excludeId = null) {
    const templates = await getAllTemplates();
    return templates.some(t => t.name === name && t.id !== excludeId);
}

// ==========================================
// 记录存储操作
// ==========================================

/**
 * 保存记录
 * @param {Object} record - 记录数据
 * @returns {Promise<Object>}
 */
async function saveRecord(record) {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.RECORDS], 'readwrite');
        const store = transaction.objectStore(STORES.RECORDS);
        
        const now = getTimestamp();
        
        const recordData = {
            ...record,
            id: record.id || generateId(),
            createdAt: record.createdAt || now,
            updatedAt: now
        };

        const request = store.put(recordData);

        request.onsuccess = () => {
            resolve(recordData);
        };

        request.onerror = () => {
            reject(new Error('保存记录失败'));
        };
    });
}

/**
 * 获取所有记录
 * @returns {Promise<Array>}
 */
async function getAllRecords() {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.RECORDS], 'readonly');
        const store = transaction.objectStore(STORES.RECORDS);
        const request = store.getAll();

        request.onsuccess = () => {
            // 按创建时间倒序排列
            const records = request.result.sort((a, b) => 
                new Date(b.createdAt) - new Date(a.createdAt)
            );
            resolve(records);
        };

        request.onerror = () => {
            reject(new Error('获取记录列表失败'));
        };
    });
}

/**
 * 根据模板 ID 获取记录
 * @param {string} templateId - 模板 ID
 * @returns {Promise<Array>}
 */
async function getRecordsByTemplateId(templateId) {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.RECORDS], 'readonly');
        const store = transaction.objectStore(STORES.RECORDS);
        const index = store.index('templateId');
        const request = index.getAll(templateId);

        request.onsuccess = () => {
            // 按创建时间倒序排列
            const records = request.result.sort((a, b) => 
                new Date(b.createdAt) - new Date(a.createdAt)
            );
            resolve(records);
        };

        request.onerror = () => {
            reject(new Error('获取记录失败'));
        };
    });
}

/**
 * 根据 ID 获取记录
 * @param {string} id - 记录 ID
 * @returns {Promise<Object|null>}
 */
async function getRecordById(id) {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.RECORDS], 'readonly');
        const store = transaction.objectStore(STORES.RECORDS);
        const request = store.get(id);

        request.onsuccess = () => {
            resolve(request.result || null);
        };

        request.onerror = () => {
            reject(new Error('获取记录失败'));
        };
    });
}

/**
 * 删除记录
 * @param {string} id - 记录 ID
 * @returns {Promise<void>}
 */
async function deleteRecord(id) {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.RECORDS], 'readwrite');
        const store = transaction.objectStore(STORES.RECORDS);
        const request = store.delete(id);

        request.onsuccess = () => {
            resolve();
        };

        request.onerror = () => {
            reject(new Error('删除记录失败'));
        };
    });
}

/**
 * 删除模板下的所有记录
 * @param {string} templateId - 模板 ID
 * @returns {Promise<void>}
 */
async function deleteRecordsByTemplateId(templateId) {
    const records = await getRecordsByTemplateId(templateId);
    
    for (const record of records) {
        await deleteRecord(record.id);
    }
}

/**
 * 根据日期范围筛选记录
 * @param {string} templateId - 模板 ID（可选）
 * @param {string} startDate - 开始日期
 * @param {string} endDate - 结束日期
 * @returns {Promise<Array>}
 */
async function getRecordsByDateRange(templateId = null, startDate = null, endDate = null) {
    let records = templateId 
        ? await getRecordsByTemplateId(templateId)
        : await getAllRecords();
    
    if (startDate) {
        const start = new Date(startDate);
        start.setHours(0, 0, 0, 0);
        records = records.filter(r => new Date(r.createdAt) >= start);
    }
    
    if (endDate) {
        const end = new Date(endDate);
        end.setHours(23, 59, 59, 999);
        records = records.filter(r => new Date(r.createdAt) <= end);
    }
    
    return records;
}

/**
 * 获取模板的记录数量
 * @param {string} templateId - 模板 ID
 * @returns {Promise<number>}
 */
async function getRecordCountByTemplateId(templateId) {
    const records = await getRecordsByTemplateId(templateId);
    return records.length;
}

// ==========================================
// 批量操作
// ==========================================

/**
 * 批量导入模板
 * @param {Array} templates - 模板数组
 * @returns {Promise<Array>}
 */
async function importTemplates(templates) {
    const results = [];
    for (const template of templates) {
        const saved = await saveTemplate(template);
        results.push(saved);
    }
    return results;
}

/**
 * 批量导入记录
 * @param {Array} records - 记录数组
 * @returns {Promise<Array>}
 */
async function importRecords(records) {
    const results = [];
    for (const record of records) {
        // 检查是否已存在相同 ID 的记录
        const existing = await getRecordById(record.id);
        if (!existing) {
            const saved = await saveRecord(record);
            results.push(saved);
        }
    }
    return results;
}

/**
 * 清空所有数据（谨慎使用）
 * @returns {Promise<void>}
 */
async function clearAllData() {
    const database = await initDB();
    
    return new Promise((resolve, reject) => {
        const transaction = database.transaction([STORES.TEMPLATES, STORES.RECORDS], 'readwrite');
        
        transaction.objectStore(STORES.TEMPLATES).clear();
        transaction.objectStore(STORES.RECORDS).clear();
        
        transaction.oncomplete = () => {
            resolve();
        };
        
        transaction.onerror = () => {
            reject(new Error('清空数据失败'));
        };
    });
}

// ==========================================
// 本地文件自动备份功能
// ==========================================

const LOCAL_BACKUP_KEY = 'wordtool_local_backup_enabled';
const LOCAL_BACKUP_DATA_KEY = 'wordtool_backup_data';

/**
 * 启用本地文件自动备份
 */
async function enableLocalBackup() {
    localStorage.setItem(LOCAL_BACKUP_KEY, 'true');
    await syncToLocalStorage();
}

/**
 * 禁用本地文件自动备份
 */
function disableLocalBackup() {
    localStorage.setItem(LOCAL_BACKUP_KEY, 'false');
}

/**
 * 检查是否启用本地备份
 */
function isLocalBackupEnabled() {
    return localStorage.getItem(LOCAL_BACKUP_KEY) === 'true';
}

/**
 * 同步数据到 localStorage（作为本地文件备份）
 */
async function syncToLocalStorage() {
    try {
        const templates = await getAllTemplates();
        const records = await getAllRecords();
        
        const backupData = {
            version: 1,
            timestamp: getTimestamp(),
            templates,
            records
        };
        
        localStorage.setItem(LOCAL_BACKUP_DATA_KEY, JSON.stringify(backupData));
        console.log('✓ 数据已自动备份到本地');
        return true;
    } catch (error) {
        console.error('自动备份失败:', error);
        return false;
    }
}

/**
 * 从 localStorage 恢复数据
 */
async function restoreFromLocalStorage() {
    try {
        const backupDataStr = localStorage.getItem(LOCAL_BACKUP_DATA_KEY);
        if (!backupDataStr) {
            console.log('没有找到本地备份');
            return false;
        }
        
        const backupData = JSON.parse(backupDataStr);
        
        // 检查是否需要恢复
        const currentTemplates = await getAllTemplates();
        if (currentTemplates.length > 0) {
            console.log('数据库已有数据，跳过自动恢复');
            return false;
        }
        
        // 恢复模板
        if (backupData.templates && backupData.templates.length > 0) {
            await importTemplates(backupData.templates, 'skip');
        }
        
        // 恢复记录
        if (backupData.records && backupData.records.length > 0) {
            await importRecords(backupData.records);
        }
        
        console.log(`✓ 已从本地备份恢复: ${backupData.templates?.length || 0} 个模板, ${backupData.records?.length || 0} 条记录`);
        return true;
    } catch (error) {
        console.error('从本地备份恢复失败:', error);
        return false;
    }
}

/**
 * 下载完整备份文件到本地
 */
async function downloadBackupFile() {
    try {
        const templates = await getAllTemplates();
        const records = await getAllRecords();
        
        const backupData = {
            version: 1,
            timestamp: getTimestamp(),
            appName: 'Word模板生成工具',
            templates,
            records
        };
        
        const blob = new Blob([JSON.stringify(backupData, null, 2)], {
            type: 'application/json'
        });
        
        const fileName = `word-tool-backup-${new Date().toISOString().slice(0, 10)}.json`;
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        return fileName;
    } catch (error) {
        console.error('下载备份文件失败:', error);
        throw error;
    }
}

/**
 * 从本地备份文件导入数据
 * @param {File} file - 备份文件
 */
async function importBackupFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                const backupData = JSON.parse(e.target.result);
                
                // 验证数据格式
                if (!backupData.templates || !backupData.records) {
                    throw new Error('备份文件格式无效');
                }
                
                // 导入模板
                let importedTemplates = 0;
                if (backupData.templates.length > 0) {
                    const result = await importTemplates(backupData.templates, 'rename');
                    importedTemplates = result.imported;
                }
                
                // 导入记录
                let importedRecords = 0;
                if (backupData.records.length > 0) {
                    const result = await importRecords(backupData.records);
                    importedRecords = result.imported;
                }
                
                // 同步到本地备份
                await syncToLocalStorage();
                
                resolve({
                    success: true,
                    imported: {
                        templates: importedTemplates,
                        records: importedRecords
                    }
                });
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => {
            reject(new Error('读取文件失败'));
        };
        
        reader.readAsText(file);
    });
}

// 导出 Storage 对象
window.Storage = {
    // 初始化
    init: initDB,
    generateId,
    getTimestamp,
    
    // 模板操作
    saveTemplate,
    getAllTemplates,
    getTemplateById,
    deleteTemplate,
    isTemplateNameExists,
    
    // 记录操作
    saveRecord,
    getAllRecords,
    getRecordsByTemplateId,
    getRecordById,
    deleteRecord,
    deleteRecordsByTemplateId,
    getRecordsByDateRange,
    getRecordCountByTemplateId,
    
    // 批量操作
    importTemplates,
    importRecords,
    clearAllData,
    
    // 本地备份功能
    enableLocalBackup,
    disableLocalBackup,
    isLocalBackupEnabled,
    syncToLocalStorage,
    restoreFromLocalStorage,
    downloadBackupFile,
    importBackupFile
};

