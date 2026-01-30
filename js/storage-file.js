/**
 * 本地文件系统存储模块
 * 使用 File System Access API 实现真正的本地文件存储
 */

// 工作区配置
const WORKSPACE_KEY = 'wordtool_workspace_handle';
let workspaceHandle = null;

/**
 * 选择工作区文件夹
 */
async function selectWorkspace() {
    try {
        // 检查浏览器是否支持 File System Access API
        if (!('showDirectoryPicker' in window)) {
            throw new Error('您的浏览器不支持本地文件系统访问，请使用 Chrome 86+ 或 Edge 86+');
        }
        
        // 让用户选择文件夹
        const dirHandle = await window.showDirectoryPicker({
            mode: 'readwrite',
            startIn: 'documents'
        });
        
        workspaceHandle = dirHandle;
        
        // 保存工作区句柄（用于下次自动加载）
        await saveWorkspaceHandle(dirHandle);
        
        // 确保存在必要的子文件夹
        await ensureWorkspaceFolders(dirHandle);
        return dirHandle;
        
    } catch (error) {
        if (error.name === 'AbortError') {
            throw new Error('用户取消选择');
        }
        throw error;
    }
}

/**
 * 保存工作区句柄到 IndexedDB（用于下次自动加载）
 */
async function saveWorkspaceHandle(dirHandle) {
    try {
        // 使用 IndexedDB 存储文件句柄
        const db = await openWorkspaceDB();
        const tx = db.transaction('workspace', 'readwrite');
        const store = tx.objectStore('workspace');
        
        await store.put({
            id: 'current',
            handle: dirHandle,
            name: dirHandle.name,
            timestamp: new Date().toISOString()
        });
        
        await tx.complete;
    } catch (error) {
        console.error('保存工作区句柄失败:', error);
    }
}

/**
 * 加载工作区句柄
 */
async function loadWorkspaceHandle() {
    try {
        const db = await openWorkspaceDB();
        const tx = db.transaction('workspace', 'readonly');
        const store = tx.objectStore('workspace');
        const result = await store.get('current');
        
        if (result && result.handle) {
            // 请求权限
            const permission = await result.handle.queryPermission({ mode: 'readwrite' });
            if (permission === 'granted' || permission === 'prompt') {
                if (permission === 'prompt') {
                    await result.handle.requestPermission({ mode: 'readwrite' });
                }
                workspaceHandle = result.handle;
                return result.handle;
            }
        }
        
        return null;
    } catch (error) {
        console.error('加载工作区句柄失败:', error);
        return null;
    }
}

/**
 * 打开工作区数据库
 */
function openWorkspaceDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open('WordToolWorkspace', 1);
        
        request.onerror = () => reject(request.error);
        request.onsuccess = () => resolve(request.result);
        
        request.onupgradeneeded = (event) => {
            const db = event.target.result;
            if (!db.objectStoreNames.contains('workspace')) {
                db.createObjectStore('workspace', { keyPath: 'id' });
            }
        };
    });
}

/**
 * 确保工作区文件夹结构存在
 */
async function ensureWorkspaceFolders(dirHandle) {
    try {
        // 创建 templates 文件夹
        await dirHandle.getDirectoryHandle('templates', { create: true });
        
        // 创建 records 文件夹
        await dirHandle.getDirectoryHandle('records', { create: true });
    } catch (error) {
        console.error('创建工作区文件夹失败:', error);
    }
}

/**
 * 获取当前工作区
 */
async function getWorkspace() {
    if (workspaceHandle) {
        return workspaceHandle;
    }
    
    // 尝试加载已保存的工作区
    const loaded = await loadWorkspaceHandle();
    if (loaded) {
        return loaded;
    }
    
    // 提示用户选择工作区
    throw new Error('未设置工作区，请先选择工作区文件夹');
}

/**
 * 检查是否已设置工作区
 */
async function hasWorkspace() {
    if (workspaceHandle) {
        return true;
    }
    
    const loaded = await loadWorkspaceHandle();
    return loaded !== null;
}

// ==========================================
// 模板文件操作
// ==========================================

/**
 * 保存模板到文件
 */
async function saveTemplateToFile(template) {
    try {
        const workspace = await getWorkspace();
        const templatesDir = await workspace.getDirectoryHandle('templates', { create: true });
        
        // 生成文件名
        const fileName = `${template.id}.json`;
        
        // 创建/获取文件
        const fileHandle = await templatesDir.getFileHandle(fileName, { create: true });
        const writable = await fileHandle.createWritable();
        
        // 写入数据
        await writable.write(JSON.stringify(template, null, 2));
        await writable.close();
        
        return template;
        
    } catch (error) {
        console.error('✗ 保存模板文件失败:', error);
        throw new Error('保存模板失败: ' + error.message);
    }
}

/**
 * 从文件加载所有模板
 */
async function loadTemplatesFromFiles() {
    try {
        const workspace = await getWorkspace();
        const templatesDir = await workspace.getDirectoryHandle('templates', { create: true });
        
        const templates = [];
        
        // 遍历文件夹中的所有文件
        for await (const entry of templatesDir.values()) {
            if (entry.kind === 'file' && entry.name.endsWith('.json')) {
                try {
                    const file = await entry.getFile();
                    const content = await file.text();
                    const template = JSON.parse(content);
                    templates.push(template);
                } catch (error) {
                    console.error(`读取模板文件失败: ${entry.name}`, error);
                }
            }
        }
        
        // 按更新时间倒序排列
        templates.sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
        
        return templates;
        
    } catch (error) {
        console.error('加载模板文件失败:', error);
        return [];
    }
}

/**
 * 从文件加载单个模板
 */
async function loadTemplateFromFile(templateId) {
    try {
        const workspace = await getWorkspace();
        const templatesDir = await workspace.getDirectoryHandle('templates', { create: true });
        
        const fileName = `${templateId}.json`;
        const fileHandle = await templatesDir.getFileHandle(fileName);
        const file = await fileHandle.getFile();
        const content = await file.text();
        
        return JSON.parse(content);
        
    } catch (error) {
        console.error('加载模板文件失败:', error);
        throw new Error('模板不存在');
    }
}

/**
 * 删除模板文件
 */
async function deleteTemplateFile(templateId) {
    try {
        const workspace = await getWorkspace();
        const templatesDir = await workspace.getDirectoryHandle('templates', { create: true });
        
        const fileName = `${templateId}.json`;
        await templatesDir.removeEntry(fileName);
        return true;
        
    } catch (error) {
        console.error('删除模板文件失败:', error);
        throw new Error('删除模板失败');
    }
}

// ==========================================
// 记录文件操作
// ==========================================

/**
 * 保存记录到文件
 */
async function saveRecordToFile(record) {
    try {
        const workspace = await getWorkspace();
        const recordsDir = await workspace.getDirectoryHandle('records', { create: true });
        
        // 按模板ID创建子文件夹
        const templateDir = await recordsDir.getDirectoryHandle(record.templateId, { create: true });
        
        // 生成文件名
        const fileName = `${record.id}.json`;
        
        // 创建/获取文件
        const fileHandle = await templateDir.getFileHandle(fileName, { create: true });
        const writable = await fileHandle.createWritable();
        
        // 写入数据
        await writable.write(JSON.stringify(record, null, 2));
        await writable.close();
        
        return record;
        
    } catch (error) {
        console.error('保存记录文件失败:', error);
        throw new Error('保存记录失败: ' + error.message);
    }
}

/**
 * 加载所有记录
 */
async function loadRecordsFromFiles(templateId = null) {
    try {
        const workspace = await getWorkspace();
        const recordsDir = await workspace.getDirectoryHandle('records', { create: true });
        
        const records = [];
        
        if (templateId) {
            // 加载特定模板的记录
            try {
                const templateDir = await recordsDir.getDirectoryHandle(templateId);
                for await (const entry of templateDir.values()) {
                    if (entry.kind === 'file' && entry.name.endsWith('.json')) {
                        try {
                            const file = await entry.getFile();
                            const content = await file.text();
                            const record = JSON.parse(content);
                            records.push(record);
                        } catch (error) {
                            console.error(`读取记录文件失败: ${entry.name}`, error);
                        }
                    }
                }
            } catch (error) {
                // 文件夹不存在，返回空数组
            }
        } else {
            // 加载所有记录
            for await (const templateEntry of recordsDir.values()) {
                if (templateEntry.kind === 'directory') {
                    for await (const entry of templateEntry.values()) {
                        if (entry.kind === 'file' && entry.name.endsWith('.json')) {
                            try {
                                const file = await entry.getFile();
                                const content = await file.text();
                                const record = JSON.parse(content);
                                records.push(record);
                            } catch (error) {
                                console.error(`读取记录文件失败: ${entry.name}`, error);
                            }
                        }
                    }
                }
            }
        }
        
        return records;
        
    } catch (error) {
        console.error('加载记录文件失败:', error);
        return [];
    }
}

/**
 * 删除记录文件
 */
async function deleteRecordFile(record) {
    try {
        const workspace = await getWorkspace();
        const recordsDir = await workspace.getDirectoryHandle('records', { create: true });
        const templateDir = await recordsDir.getDirectoryHandle(record.templateId);
        
        const fileName = `${record.id}.json`;
        await templateDir.removeEntry(fileName);
        return true;
        
    } catch (error) {
        console.error('删除记录文件失败:', error);
        throw new Error('删除记录失败');
    }
}

/**
 * 删除模板的所有记录文件
 */
async function deleteRecordsByTemplateId(templateId) {
    try {
        const workspace = await getWorkspace();
        const recordsDir = await workspace.getDirectoryHandle('records', { create: true });
        
        await recordsDir.removeEntry(templateId, { recursive: true });
        return true;
        
    } catch (error) {
        if (error.name === 'NotFoundError') {
            return true; // 文件夹不存在，视为成功
        }
        console.error('删除记录文件夹失败:', error);
        throw new Error('删除记录失败');
    }
}

// 导出
window.FileStorage = {
    // 工作区管理
    selectWorkspace,
    getWorkspace,
    hasWorkspace,
    loadWorkspaceHandle,
    
    // 模板操作
    saveTemplateToFile,
    loadTemplatesFromFiles,
    loadTemplateFromFile,
    deleteTemplateFile,
    
    // 记录操作
    saveRecordToFile,
    loadRecordsFromFiles,
    deleteRecordFile,
    deleteRecordsByTemplateId
};
