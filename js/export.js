/**
 * 导入导出模块
 * 提供模板包和数据包的导入导出功能
 */

// 文件扩展名
const FILE_EXTENSIONS = {
    TEMPLATE_PACK: '.templatepack',
    RECORDS_PACK: '.recordspack'
};

// 包版本号（用于未来兼容性处理）
const PACK_VERSION = '1.0.0';

/**
 * 下载文件（原生方式，更可靠）
 * @param {Blob} blob - 文件 Blob
 * @param {string} fileName - 文件名
 */
function downloadFile(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 100);
}

/**
 * 创建模板包
 * @param {Object} template - 模板数据
 * @returns {Object} - 模板包数据
 */
function createTemplatePack(template) {
    return {
        version: PACK_VERSION,
        type: 'template',
        exportedAt: new Date().toISOString(),
        template: {
            id: template.id,
            name: template.name,
            description: template.description,
            titleFieldKey: template.titleFieldKey,
            wordFile: template.wordFile,
            fields: template.fields,
            detailTables: template.detailTables,
            createdAt: template.createdAt,
            updatedAt: template.updatedAt
        }
    };
}

/**
 * 导出模板包
 * @param {string} templateId - 模板 ID
 * @returns {Promise<void>}
 */
async function exportTemplatePack(templateId) {
    const template = await StorageAdapter.getTemplateById(templateId);
    
    if (!template) {
        throw new Error('模板不存在');
    }
    
    const pack = createTemplatePack(template);
    const json = JSON.stringify(pack, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    
    // 生成文件名
    const fileName = `${template.name}_模板包_${Word.formatDate(new Date(), 'YYYYMMDD')}${FILE_EXTENSIONS.TEMPLATE_PACK}`;
    
    downloadFile(blob, fileName);
}

/**
 * 导入模板包
 * @param {File} file - 模板包文件
 * @param {Object} options - 导入选项
 * @param {string} options.conflictStrategy - 冲突策略: 'overwrite' | 'rename' | 'skip'
 * @returns {Promise<Object>} - 导入结果
 */
async function importTemplatePack(file, options = { conflictStrategy: 'rename' }) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                const content = e.target.result;
                const pack = JSON.parse(content);
                
                // 验证包格式
                if (!pack.type || pack.type !== 'template' || !pack.template) {
                    throw new Error('无效的模板包格式');
                }
                
                const template = pack.template;
                
                // 检查是否存在同名模板
                const existingTemplates = await StorageAdapter.getAllTemplates();
                const existingByName = existingTemplates.find(t => t.name === template.name);
                const existingById = existingTemplates.find(t => t.id === template.id);
                
                let finalTemplate = { ...template };
                let action = 'created';
                
                if (existingById || existingByName) {
                    switch (options.conflictStrategy) {
                        case 'overwrite':
                            // 使用现有的 ID 进行覆盖
                            finalTemplate.id = existingById?.id || existingByName?.id;
                            action = 'overwritten';
                            break;
                        case 'rename':
                            // 生成新的 ID 和名称
                            finalTemplate.id = StorageAdapter.generateId();
                            finalTemplate.name = `${template.name} (导入-${Word.formatDate(new Date(), 'MMDDHHmm')})`;
                            finalTemplate.createdAt = null;
                            action = 'renamed';
                            break;
                        case 'skip':
                            resolve({
                                success: true,
                                action: 'skipped',
                                message: `模板 "${template.name}" 已存在，跳过导入`
                            });
                            return;
                    }
                } else {
                    // 生成新 ID
                    finalTemplate.id = StorageAdapter.generateId();
                    finalTemplate.createdAt = null;
                }
                
                // 保存模板
                const saved = await StorageAdapter.saveTemplate(finalTemplate);
                
                resolve({
                    success: true,
                    action,
                    template: saved,
                    message: `模板 "${saved.name}" 导入成功`
                });
                
            } catch (error) {
                reject(new Error(`导入模板包失败: ${error.message}`));
            }
        };
        
        reader.onerror = () => {
            reject(new Error('读取文件失败'));
        };
        
        reader.readAsText(file);
    });
}

/**
 * 创建数据包
 * @param {string} templateId - 模板 ID
 * @param {Array} records - 记录列表
 * @param {Object} template - 模板数据（用于包含模板名称等信息）
 * @returns {Object} - 数据包数据
 */
function createRecordsPack(templateId, records, template) {
    return {
        version: PACK_VERSION,
        type: 'records',
        exportedAt: new Date().toISOString(),
        templateInfo: {
            id: templateId,
            name: template?.name || '未知模板'
        },
        recordCount: records.length,
        records: records.map(record => ({
            id: record.id,
            templateId: record.templateId,
            data: record.data,
            tables: record.tables,
            createdAt: record.createdAt,
            updatedAt: record.updatedAt
        }))
    };
}

/**
 * 导出数据包
 * @param {string} templateId - 模板 ID
 * @param {string} startDate - 开始日期（可选）
 * @param {string} endDate - 结束日期（可选）
 * @returns {Promise<void>}
 */
async function exportRecordsPack(templateId, startDate = null, endDate = null) {
    const template = await StorageAdapter.getTemplateById(templateId);
    
    if (!template) {
        throw new Error('模板不存在');
    }
    
    // 获取指定日期范围的记录
    const records = await StorageAdapter.getRecordsByDateRange(templateId, startDate, endDate);
    
    if (records.length === 0) {
        throw new Error('没有找到符合条件的记录');
    }
    
    const pack = createRecordsPack(templateId, records, template);
    const json = JSON.stringify(pack, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    
    // 生成文件名
    let dateRange = '';
    if (startDate || endDate) {
        const start = startDate ? Word.formatDate(new Date(startDate), 'YYYYMMDD') : '起始';
        const end = endDate ? Word.formatDate(new Date(endDate), 'YYYYMMDD') : '至今';
        dateRange = `_${start}_至_${end}`;
    }
    
    const fileName = `${template.name}_数据包${dateRange}_${Word.formatDate(new Date(), 'YYYYMMDD')}${FILE_EXTENSIONS.RECORDS_PACK}`;
    
    downloadFile(blob, fileName);
}

/**
 * 导入数据包
 * @param {File} file - 数据包文件
 * @returns {Promise<Object>} - 导入结果
 */
async function importRecordsPack(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                const content = e.target.result;
                const pack = JSON.parse(content);
                
                // 验证包格式
                if (!pack.type || pack.type !== 'records' || !pack.records) {
                    throw new Error('无效的数据包格式');
                }
                
                // 检查对应的模板是否存在
                const templateId = pack.templateInfo?.id;
                let targetTemplateId = templateId;
                
                if (templateId) {
                    const template = await StorageAdapter.getTemplateById(templateId);
                    if (!template) {
                        // 尝试按名称查找模板
                        const templates = await StorageAdapter.getAllTemplates();
                        const templateByName = templates.find(t => t.name === pack.templateInfo?.name);
                        if (templateByName) {
                            targetTemplateId = templateByName.id;
                        } else {
                            throw new Error(`找不到对应的模板 "${pack.templateInfo?.name || templateId}"，请先导入模板`);
                        }
                    }
                }
                
                // 导入记录
                let imported = 0;
                let skipped = 0;
                
                for (const record of pack.records) {
                    // 检查是否已存在
                    const existing = await StorageAdapter.getRecordById(record.id);
                    
                    if (existing) {
                        skipped++;
                        continue;
                    }
                    
                    // 更新模板 ID（如果模板 ID 发生了变化）
                    const recordToSave = {
                        ...record,
                        templateId: targetTemplateId,
                        id: StorageAdapter.generateId() // 生成新 ID 避免冲突
                    };
                    
                    await StorageAdapter.saveRecord(recordToSave);
                    imported++;
                }
                
                resolve({
                    success: true,
                    imported,
                    skipped,
                    total: pack.records.length,
                    message: `成功导入 ${imported} 条记录${skipped > 0 ? `，跳过 ${skipped} 条重复记录` : ''}`
                });
                
            } catch (error) {
                reject(new Error(`导入数据包失败: ${error.message}`));
            }
        };
        
        reader.onerror = () => {
            reject(new Error('读取文件失败'));
        };
        
        reader.readAsText(file);
    });
}

/**
 * 导出所有数据（完整备份）
 * @returns {Promise<void>}
 */
async function exportFullBackup() {
    const templates = await StorageAdapter.getAllTemplates();
    const records = await StorageAdapter.getAllRecords();
    
    const backup = {
        version: PACK_VERSION,
        type: 'full_backup',
        exportedAt: new Date().toISOString(),
        templates,
        records
    };
    
    const json = JSON.stringify(backup, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    
    const fileName = `Word模板工具_完整备份_${Word.formatDate(new Date(), 'YYYYMMDD_HHmmss')}.backup`;
    
    downloadFile(blob, fileName);
}

/**
 * 导入完整备份
 * @param {File} file - 备份文件
 * @returns {Promise<Object>} - 导入结果
 */
async function importFullBackup(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = async (e) => {
            try {
                const content = e.target.result;
                const backup = JSON.parse(content);
                
                // 验证备份格式
                if (!backup.type || backup.type !== 'full_backup') {
                    throw new Error('无效的备份文件格式');
                }
                
                let templatesImported = 0;
                let recordsImported = 0;
                
                // 导入模板
                if (backup.templates && backup.templates.length > 0) {
                    for (const template of backup.templates) {
                        const existing = await StorageAdapter.getTemplateById(template.id);
                        if (!existing) {
                            await StorageAdapter.saveTemplate(template);
                            templatesImported++;
                        }
                    }
                }
                
                // 导入记录
                if (backup.records && backup.records.length > 0) {
                    for (const record of backup.records) {
                        const existing = await StorageAdapter.getRecordById(record.id);
                        if (!existing) {
                            await StorageAdapter.saveRecord(record);
                            recordsImported++;
                        }
                    }
                }
                
                resolve({
                    success: true,
                    templatesImported,
                    recordsImported,
                    message: `成功导入 ${templatesImported} 个模板，${recordsImported} 条记录`
                });
                
            } catch (error) {
                reject(new Error(`导入备份失败: ${error.message}`));
            }
        };
        
        reader.onerror = () => {
            reject(new Error('读取文件失败'));
        };
        
        reader.readAsText(file);
    });
}

/**
 * 读取文件内容
 * @param {File} file - 文件对象
 * @returns {Promise<string>} - 文件内容
 */
function readFileContent(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            resolve(e.target.result);
        };
        
        reader.onerror = () => {
            reject(new Error('读取文件失败'));
        };
        
        reader.readAsText(file);
    });
}

/**
 * 验证文件类型
 * @param {File} file - 文件对象
 * @param {string} expectedExtension - 期望的扩展名
 * @returns {boolean}
 */
function validateFileType(file, expectedExtension) {
    const fileName = file.name.toLowerCase();
    return fileName.endsWith(expectedExtension.toLowerCase());
}

// 导出 Export 对象
window.Export = {
    FILE_EXTENSIONS,
    PACK_VERSION,
    
    // 模板包
    createTemplatePack,
    exportTemplatePack,
    importTemplatePack,
    
    // 数据包
    createRecordsPack,
    exportRecordsPack,
    importRecordsPack,
    
    // 完整备份
    exportFullBackup,
    importFullBackup,
    
    // 工具函数
    readFileContent,
    validateFileType
};

