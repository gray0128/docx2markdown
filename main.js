// 文件系统和路径处理
const fs = require('fs');
const path = require('path');
const pypandoc = require('pypandoc');

// 递归地获取文件夹内的所有文件
function getFiles(folderPath, recursive) {
    let files = [];
    if (recursive) {
        const walkSync = (dir, filelist = []) => {
            fs.readdirSync(dir).forEach(file => {
                const filePath = path.join(dir, file);
                if (fs.statSync(filePath).isDirectory()) {
                    filelist = walkSync(filePath, filelist);
                } else {
                    filelist.push(filePath);
                }
            });
            return filelist;
        };
        files = walkSync(folderPath);
    } else {
        files = fs.readdirSync(folderPath).map(file => path.join(folderPath, file));
    }
    return files.filter(file => path.extname(file) === '.docx');
}

// DOCX to Markdown 转换
async function convertDocxToMd(folderPath, recursive, mediaExtraction) {
    const docxFiles = getFiles(folderPath, recursive);
    const totalFiles = docxFiles.length;
    const mediaFolder = path.join(folderPath, "media");
    
    if (mediaExtraction) {
        fs.mkdirSync(mediaFolder, { recursive: true });
    }

    for (let i = 0; i < totalFiles; i++) {
        const docxFile = docxFiles[i];
        const mdPath = path.join(path.dirname(docxFile), path.basename(docxFile, '.docx') + '.md');
        
        let extraArgs = ['--wrap=none'];
        if (mediaExtraction) {
            extraArgs.push(`--extract-media=${mediaFolder}`);
        }

        await new Promise((resolve) => {
            pypandoc.convertFile(docxFile, 'md', { outputfile: mdPath, extraArgs }, (err) => {
                if (err) {
                    console.error(err);
                }
                resolve();
            });
        });

        updateProgress((i + 1) / totalFiles * 100);
    }
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
    const folderPath = document.getElementById('folderPath').files[0].path;
    const recursive = document.getElementById('recursive').checked;
    const mediaExtraction = document.getElementById('mediaExtraction').checked;

    try {
        await convertDocxToMd(folderPath, recursive, mediaExtraction);
        updateProgress(100);
        document.getElementById('percentageLabel').textContent = 'Done!';
    } catch (err) {
        console.error('Conversion failed:', err);
        document.getElementById('percentageLabel').textContent = 'Error!';
    }
});
