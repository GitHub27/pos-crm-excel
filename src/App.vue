<script setup>
import { ref } from "vue";
import { invoke } from "@tauri-apps/api/core";
import { open, save } from '@tauri-apps/plugin-dialog';

const greetMsg = ref("");
const name = ref("");
const selectedFile = ref(null);
const columnName = ref("省市区");
const processing = ref(false);
const result = ref(null);

async function greet() {
  greetMsg.value = await invoke("greet", { name: name.value });
}

async function selectFile() {
  try {
    const selected = await open({
      multiple: false,
      filters: [{
        name: 'Excel Files',
        extensions: ['xlsx', 'xls']
      }]
    });
    
    if (selected) {
      selectedFile.value = selected;
    }
  } catch (error) {
    console.error('选择文件时出错:', error);
  }
}

async function processFile() {
  if (!selectedFile.value) {
    alert('请先选择一个Excel文件');
    return;
  }
  
  if (!columnName.value.trim()) {
    alert('请输入列名');
    return;
  }
  
  try {
    processing.value = true;
    result.value = null;
    
    // 生成当天日期作为默认文件名
    const today = new Date();
    const dateStr = today.getFullYear() + '-' + 
                   String(today.getMonth() + 1).padStart(2, '0') + '-' + 
                   String(today.getDate()).padStart(2, '0');
    const defaultFileName = `${dateStr}.xlsx`;
    
    // 选择保存位置
    const outputPath = await save({
      defaultPath: defaultFileName,
      filters: [{
        name: 'Excel Files',
        extensions: ['xlsx']
      }]
    });
    
    if (!outputPath) {
      processing.value = false;
      return;
    }
    
    // 生成当前ISO格式日期
    const currentDate = new Date().toISOString();
    
    // 调用Rust命令处理Excel文件
    const processResult = await invoke('process_excel', {
      inputPath: selectedFile.value,
      outputPath: outputPath,
      columnName: columnName.value.trim(),
      createDate: currentDate,
      updateDate: currentDate
    });
    
    result.value = processResult;
    
    if (processResult.success) {
      alert('文件处理成功！已保存到: ' + outputPath);
    } else {
      alert('处理失败: ' + processResult.message);
    }
  } catch (error) {
    console.error('处理文件时出错:', error);
    alert('处理文件时出错: ' + error);
  } finally {
    processing.value = false;
  }
}

function resetForm() {
  selectedFile.value = null;
  columnName.value = "省市区";
  result.value = null;
}
</script>

<template>
  <main class="container">
    <h1>Excel 文件处理工具</h1>
    
    <!-- Excel文件处理区域 -->
    <div class="excel-processor">
      <h2>上传并处理Excel文件</h2>
      
      <div class="form-group">
        <label>选择Excel文件:</label>
        <div class="file-input-group">
          <button @click="selectFile" class="btn btn-secondary">
            {{ selectedFile ? '重新选择文件' : '选择文件' }}
          </button>
          <span v-if="selectedFile" class="file-name">{{ selectedFile }}</span>
        </div>
      </div>
      
      <!-- <div class="form-group">
        <label for="column-name">目标列名:</label>
        <input 
          id="column-name" 
          v-model="columnName" 
          placeholder="请输入列名，如：省市区" 
          class="input"
        />
      </div> -->
      
      <div class="button-group">
        <button 
          @click="processFile" 
          :disabled="!selectedFile || processing" 
          class="btn btn-primary"
        >
          {{ processing ? '处理中...' : '处理文件' }}
        </button>
        <button @click="resetForm" class="btn btn-secondary">重置</button>
      </div>
      
      <!-- 处理结果 -->
      <div v-if="result" class="result">
        <div v-if="result.success" class="success">
          <h3>✅ 处理成功</h3>
          <p>{{ result.message }}</p>
          <p v-if="result.output_path">保存位置: {{ result.output_path }}</p>
        </div>
        <div v-else class="error">
          <h3>❌ 处理失败</h3>
          <p>{{ result.message }}</p>
        </div>
      </div>
    </div>
    
    <!-- 原有的问候功能 -->
    <!-- <div class="divider"></div>
    
    <div class="greet-section">
      <h2>问候功能</h2>
      <form class="row" @submit.prevent="greet">
        <input id="greet-input" v-model="name" placeholder="Enter a name..." class="input" />
        <button type="submit" class="btn btn-primary">Greet</button>
      </form>
      <p v-if="greetMsg" class="greet-message">{{ greetMsg }}</p>
    </div> -->
  </main>
</template>

<style scoped>
/* Excel处理器样式 */
.excel-processor {
  background: white;
  border-radius: 12px;
  padding: 2rem;
  margin: 2rem 0;
  box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
  text-align: left;
}

.excel-processor h2 {
  color: #2c3e50;
  margin-bottom: 1.5rem;
  text-align: center;
}

.form-group {
  margin-bottom: 1.5rem;
}

.form-group label {
  display: block;
  margin-bottom: 0.5rem;
  font-weight: 600;
  color: #34495e;
}

.file-input-group {
  display: flex;
  align-items: center;
  gap: 1rem;
}

.file-name {
  font-size: 0.9rem;
  color: #7f8c8d;
  word-break: break-all;
}

.input {
  width: 100%;
  padding: 0.75rem;
  border: 2px solid #e1e8ed;
  border-radius: 8px;
  font-size: 1rem;
  transition: border-color 0.3s;
}

.input:focus {
  outline: none;
  border-color: #3498db;
}

.btn {
  padding: 0.75rem 1.5rem;
  border: none;
  border-radius: 8px;
  font-size: 1rem;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.3s;
}

.btn:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.btn-primary {
  background: #3498db;
  color: white;
}

.btn-primary:hover:not(:disabled) {
  background: #2980b9;
}

.btn-secondary {
  background: #95a5a6;
  color: white;
}

.btn-secondary:hover {
  background: #7f8c8d;
}

.button-group {
  display: flex;
  gap: 1rem;
  justify-content: center;
}

.result {
  margin-top: 2rem;
  padding: 1rem;
  border-radius: 8px;
}

.success {
  background: #d4edda;
  border: 1px solid #c3e6cb;
  color: #155724;
}

.error {
  background: #f8d7da;
  border: 1px solid #f5c6cb;
  color: #721c24;
}

.result h3 {
  margin: 0 0 0.5rem 0;
}

.result p {
  margin: 0.25rem 0;
}

.divider {
  height: 2px;
  background: linear-gradient(to right, transparent, #e1e8ed, transparent);
  margin: 3rem 0;
}

.greet-section {
  background: white;
  border-radius: 12px;
  padding: 2rem;
  box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}

.greet-section h2 {
  color: #2c3e50;
  margin-bottom: 1.5rem;
}

.greet-message {
  margin-top: 1rem;
  padding: 1rem;
  background: #e8f4fd;
  border-radius: 8px;
  color: #2c3e50;
}

</style>
<style>
:root {
  font-family: Inter, Avenir, Helvetica, Arial, sans-serif;
  font-size: 16px;
  line-height: 24px;
  font-weight: 400;

  color: #0f0f0f;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  min-height: 100vh;

  font-synthesis: none;
  text-rendering: optimizeLegibility;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  -webkit-text-size-adjust: 100%;
}

.container {
  max-width: 800px;
  margin: 0 auto;
  padding: 2rem;
  min-height: 100vh;
}

.container h1 {
  text-align: center;
  color: white;
  font-size: 2.5rem;
  margin-bottom: 2rem;
  text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
}

.row {
  display: flex;
  gap: 1rem;
  align-items: center;
}

a {
  font-weight: 500;
  color: #646cff;
  text-decoration: inherit;
}

a:hover {
  color: #535bf2;
}

h1 {
  text-align: center;
}

input,
button {
  border-radius: 8px;
  border: 1px solid transparent;
  padding: 0.6em 1.2em;
  font-size: 1em;
  font-weight: 500;
  font-family: inherit;
  color: #0f0f0f;
  background-color: #ffffff;
  transition: border-color 0.25s;
  box-shadow: 0 2px 2px rgba(0, 0, 0, 0.2);
}

button {
  cursor: pointer;
}

button:hover {
  border-color: #396cd8;
}
button:active {
  border-color: #396cd8;
  background-color: #e8e8e8;
}

input,
button {
  outline: none;
}

#greet-input {
  margin-right: 5px;
}

@media (prefers-color-scheme: dark) {
  :root {
    color: #f6f6f6;
    background-color: #2f2f2f;
  }

  a:hover {
    color: #24c8db;
  }

  input,
  button {
    color: #ffffff;
    background-color: #0f0f0f98;
  }
  button:active {
    background-color: #0f0f0f69;
  }
}

</style>
