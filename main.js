const fs = require('fs');
const path = require('path');
const pypandoc = require('pypandoc');

// DOCX to Markdown 转换
async function convertDocxToMd(filePath, mediaExtraction) {
    const mediaFolder = path.join(path.dirname(filePath), "media");
    
    if (mediaExtraction) {
        fs.mkdirSync(mediaFolder, { recursive: true });
    }

    const mdPath = path.join(path.dirname(filePath), path.basename(filePath, '.docx') + '.md');
    
    let extraArgs = ['--wrap=none'];
    if (mediaExtraction) {
        extraArgs.push(`--extract-media=${mediaFolder}`);
    }

    await new Promise((resolve, reject) => {
        pypandoc.convertFile(filePath, 'md', { outputfile: mdPath, extraArgs }, (err) => {
            if (err) {
                reject(err);
            } else {
                resolve();
            }
        });
    });

    // 使用 requestAnimationFrame 确保UI更新在主线程上执行
    requestAnimationFrame(() => {
        updateProgress(100);
        document.getElementById('percentageLabel').textContent = 'Done!';
    });
}

// 更新进度条和百分比
function updateProgress(value) {
    const progressBar = document.getElementById('progress');
    const percentageLabel = document.getElementById('percentageLabel');

    // 使用 requestAnimationFrame 确保UI更新在主线程上执行
    requestAnimationFrame(() => {
        progressBar.style.width = `${value}%`;
        percentageLabel.textContent = `${Math.round(value)}%`;
    });
}

// 事件处理
document.getElementById('convertBtn').addEventListener('click', async () => {
    const fileInput = document.getElementById('fileInput');
    const mediaExtraction = document.getElementById('mediaExtraction').checked;

    if (!fileInput.files.length) {
        alert('Please select a DOCX file.');
        return;
    }

    const filePath = fileInput.files[0].path;
    
    try {
        // 开始转换前更新进度条
        updateProgress(0);
        await convertDocxToMd(filePath, mediaExtraction);
    } catch (err) {
        console.error('Conversion failed:', err);
        // 使用 requestAnimationFrame 确保错误信息更新在主线程上执行
        requestAnimationFrame(() => {
            document.getElementById('percentageLabel').textContent = 'Error!';
        });
    }
});
