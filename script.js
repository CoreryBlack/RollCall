// 初始化
let names = JSON.parse(localStorage.getItem('roll_names') || '[]');
let picked = new Set(JSON.parse(localStorage.getItem('roll_picked') || '[]'));

const namesArea = document.getElementById('names');
const resultDiv = document.getElementById('result');
const pickedList = document.getElementById('pickedList');

function render() {
    namesArea.value = names.join('\n');
    pickedList.innerHTML = '';
    [...picked].forEach(n => {
        const li = document.createElement('li');
        li.textContent = n;
        pickedList.appendChild(li);
    });
}

function saveState() {
    localStorage.setItem('roll_names', JSON.stringify(names));
    localStorage.setItem('roll_picked', JSON.stringify([...picked]));
}

function candidates() {
    return names.filter(n => !picked.has(n));
}

function randomPick(list) {
    const arr = new Uint32Array(1);
    crypto.getRandomValues(arr);
    return list[arr[0] % list.length];
}

// 抽取
document.getElementById('draw').onclick = () => {
    const pool = candidates();
    if (pool.length === 0) {
        resultDiv.textContent = '本周期已全部覆盖，请重置';
        return;
    }

    resultDiv.classList.add('drawing');

    const animationDuration = 1500;
    const frameRate = 50;
    const totalFrames = animationDuration / frameRate;
    let currentFrame = 0;

    const animateInterval = setInterval(() => {
        const randomCandidate = randomPick(pool);
        resultDiv.textContent = `抽中：${randomCandidate}`;
        currentFrame++;

        if (currentFrame >= totalFrames) {
            clearInterval(animateInterval);
            const finalName = randomPick(pool);
            picked.add(finalName);
            saveState();
            render();
            resultDiv.textContent = `抽中：${finalName}`;
            resultDiv.classList.remove('drawing');
            resultDiv.classList.add('draw-complete');
            setTimeout(() => {
                resultDiv.classList.remove('draw-complete');
            }, 1000);
        }
    }, frameRate);
};

// 重置
document.getElementById('reset').onclick = () => {
    picked.clear();
    saveState();
    render();
    resultDiv.textContent = '已重置本周期';
};

// 保存名单
document.getElementById('saveNames').onclick = () => {
    names = namesArea.value.split('\n').map(s => s.trim()).filter(Boolean);
    saveState();
    render();
    alert('名单已保存');
};

// 导出
document.getElementById('exportJson').onclick = () => {
    const data = { names, picked: [...picked] };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'rollcall.json';
    a.click();
};

// 导入JSON
document.getElementById('importJson').onclick = () => {
    document.getElementById('fileInput').click();
};
document.getElementById('fileInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    const text = await file.text();
    const cfg = JSON.parse(text);
    if (Array.isArray(cfg.names)) names = cfg.names;
    if (Array.isArray(cfg.picked)) picked = new Set(cfg.picked);
    saveState();
    render();
    alert('导入成功');
    e.target.value = '';
});

// 导入Excel
document.getElementById('importExcel').onclick = () => {
    document.getElementById('excelInput').click();
};
document.getElementById('excelInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                // 读取为二维数组
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (rows.length < 2) {
                    alert('Excel文件中没有找到有效表头');
                    return;
                }

                // 第二行才是真正表头
                const headers = rows[1].map(h => String(h || '').trim());

                // 支持的姓名列别名
                const nameAliases = ['姓名', '名字', 'Name', 'name'];

                // 查找姓名列索引
                const nameIndex = headers.findIndex(h => nameAliases.includes(h));
                if (nameIndex === -1) {
                    alert('未找到姓名列，请检查Excel表头是否为“姓名”或常见别名');
                    return;
                }

                // 从第三行开始才是数据
                const dataRows = rows.slice(2);

                const extractedNames = dataRows.map(row => {
                    const name = row[nameIndex];
                    return typeof name === 'string' ? name.trim() : String(name || '').trim();
                }).filter(name => name.length > 0);

                if (extractedNames.length === 0) {
                    alert('未能从姓名列提取到有效名字');
                    return;
                }

                names = extractedNames;
                saveState();
                render();
                alert(`成功导入 ${extractedNames.length} 个名字`);
            } catch (error) {
                console.error('Excel解析错误:', error);
                alert('Excel文件解析失败: ' + error.message);
            }
        };

        reader.onerror = function() {
            alert('文件读取失败');
        };

        reader.readAsArrayBuffer(file);
    } catch (error) {
        console.error('Excel导入错误:', error);
        alert('Excel导入失败: ' + error.message);
    } finally {
        e.target.value = '';
    }
});

// 页面加载时渲染
render();