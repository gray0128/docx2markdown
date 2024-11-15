// 文件系统和路径处理
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

    updateProgress(100);
}

// 更新进度条和百分比
function updateProgress(value) {
    const progressBar = document.getElementById('progress');
    const percentageLabel = document.getElementById('percentageLabel');

    progressBar.style.width = `${value}%`;
    percentageLabel.textContent = `${Math.round(value)}%`;
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
        await convertDocxToMd(filePath, mediaExtraction);
        document.getElementById('percentageLabel').textContent = 'Done!';
    } catch (err) {
        console.error('Conversion failed:', err);
        document.getElementById('percentageLabel').textContent = 'Error!';
    }
});
