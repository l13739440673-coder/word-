/**
 * 存储适配器
 * 统一接口，支持文件存储和浏览器存储切换
 */

// 当前存储模式
let storageMode = 'file'; // 'file' 或 'browser'

/**
 * 初始化存储
 */
async function initStorage() {
    // 检查是否支持文件系统 API
    const supportsFileSystem = 'showDirectoryPicker' in window;
    
    if (supportsFileSystem) {
        // 尝试加载文件存储
        const hasWorkspace = await FileStorage.hasWorkspace();
        if (hasWorkspace) {
            storageMode = 'file';
            return;
        }
    }
    
    // 回退到浏览器存储
    storageMode = 'browser';
    await Storage.init();
}

/**
 * 获取当前存储模式
 */
function getStorageMode() {
    return storageMode;
}

/**
 * 切换到文件存储模式
 */
async function switchToFileStorage() {
    try {
        await FileStorage.selectWorkspace();
        storageMode = 'file';
        return true;
    } catch (error) {
        throw error;
    }
}

// ==========================================
// 统一的存储接口
// ==========================================

/**
 * 保存模板
 */
async function saveTemplate(template) {
    // 确保模板有ID和时间戳
    if (!template.id) {
        template.id = Storage.generateId();
    }
    if (!template.createdAt) {
        template.createdAt = Storage.getTimestamp();
    }
    template.updatedAt = Storage.getTimestamp();
    
    if (storageMode === 'file') {
        return await FileStorage.saveTemplateToFile(template);
    } else {
        return await Storage.saveTemplate(template);
    }
}

/**
 * 获取所有模板
 */
async function getAllTemplates() {
    if (storageMode === 'file') {
        return await FileStorage.loadTemplatesFromFiles();
    } else {
        return await Storage.getAllTemplates();
    }
}

/**
 * 根据ID获取模板
 */
async function getTemplateById(id) {
    if (storageMode === 'file') {
        return await FileStorage.loadTemplateFromFile(id);
    } else {
        return await Storage.getTemplateById(id);
    }
}

/**
 * 删除模板
 */
async function deleteTemplate(id) {
    if (storageMode === 'file') {
        await FileStorage.deleteTemplateFile(id);
        await FileStorage.deleteRecordsByTemplateId(id);
        return true;
    } else {
        return await Storage.deleteTemplate(id);
    }
}

/**
 * 保存记录
 */
async function saveRecord(record) {
    // 确保有ID和时间戳
    if (!record.id) {
        record.id = Storage.generateId();
    }
    if (!record.createdAt) {
        record.createdAt = Storage.getTimestamp();
    }
    record.updatedAt = Storage.getTimestamp();
    
    if (storageMode === 'file') {
        return await FileStorage.saveRecordToFile(record);
    } else {
        return await Storage.saveRecord(record);
    }
}

/**
 * 获取所有记录
 */
async function getAllRecords() {
    if (storageMode === 'file') {
        return await FileStorage.loadRecordsFromFiles();
    } else {
        return await Storage.getAllRecords();
    }
}

/**
 * 根据模板ID获取记录
 */
async function getRecordsByTemplateId(templateId) {
    if (storageMode === 'file') {
        return await FileStorage.loadRecordsFromFiles(templateId);
    } else {
        return await Storage.getRecordsByTemplateId(templateId);
    }
}

/**
 * 根据ID获取记录
 */
async function getRecordById(id) {
    if (storageMode === 'file') {
        const allRecords = await FileStorage.loadRecordsFromFiles();
        return allRecords.find(r => r.id === id);
    } else {
        return await Storage.getRecordById(id);
    }
}

/**
 * 删除记录
 */
async function deleteRecord(id) {
    if (storageMode === 'file') {
        // 先找到记录以获取templateId
        const record = await getRecordById(id);
        if (record) {
            await FileStorage.deleteRecordFile(record);
        }
        return true;
    } else {
        return await Storage.deleteRecord(id);
    }
}

/**
 * 删除模板的所有记录
 */
async function deleteRecordsByTemplateId(templateId) {
    if (storageMode === 'file') {
        return await FileStorage.deleteRecordsByTemplateId(templateId);
    } else {
        return await Storage.deleteRecordsByTemplateId(templateId);
    }
}

/**
 * 按日期范围获取记录
 */
async function getRecordsByDateRange(templateId, startDate, endDate) {
    const records = templateId ? 
        await getRecordsByTemplateId(templateId) : 
        await getAllRecords();
    
    return records.filter(record => {
        const recordDate = new Date(record.createdAt);
        const start = startDate ? new Date(startDate) : null;
        const end = endDate ? new Date(endDate + 'T23:59:59') : null;
        
        if (start && recordDate < start) return false;
        if (end && recordDate > end) return false;
        return true;
    });
}

/**
 * 获取记录数量
 */
async function getRecordCountByTemplateId(templateId) {
    const records = await getRecordsByTemplateId(templateId);
    return records.length;
}

// 导出适配器
window.StorageAdapter = {
    // 初始化
    init: initStorage,
    getStorageMode,
    switchToFileStorage,
    generateId: Storage.generateId,
    getTimestamp: Storage.getTimestamp,
    
    // 模板操作
    saveTemplate,
    getAllTemplates,
    getTemplateById,
    deleteTemplate,
    isTemplateNameExists: Storage.isTemplateNameExists,
    
    // 记录操作
    saveRecord,
    getAllRecords,
    getRecordsByTemplateId,
    getRecordById,
    deleteRecord,
    deleteRecordsByTemplateId,
    getRecordsByDateRange,
    getRecordCountByTemplateId,
    
    // 批量操作（使用浏览器存储的方法）
    importTemplates: Storage.importTemplates,
    importRecords: Storage.importRecords,
    clearAllData: Storage.clearAllData,
    
    // 备份功能（使用浏览器存储的方法）
    downloadBackupFile: Storage.downloadBackupFile,
    importBackupFile: Storage.importBackupFile
};
