function showPopup(message, type) {
    const popup = document.getElementById('popup');
    popup.textContent = message;
    popup.className = 'popup show ' + type;

    setTimeout(() => {
        popup.classList.remove('show');
    }, 3000); // 弹窗显示3秒后消失
}

function uploadExcel() {
    const fileInput = document.getElementById('file');
    const startHeaderRow = document.getElementById('startHeaderRow').value;
    const endHeaderRow = document.getElementById('endHeaderRow').value;
    const startDataRow = document.getElementById('startDataRow').value;
    const endDataRow = document.getElementById('endDataRow').value;
    const columnCount = document.getElementById('columnCount').value;
    const templateName = document.getElementById('templateName').value;

    // 客户端验证
    if (!fileInput.files.length) {
        showPopup('请选择一个文件上传。', 'error');
        return;
    }

    if (parseInt(startHeaderRow) > parseInt(endHeaderRow)) {
        showPopup('表头起始行不能大于表头结束行。', 'error');
        return;
    }

    if (parseInt(startDataRow) > parseInt(endDataRow)) {
        showPopup('数据起始行不能大于数据结束行。', 'error');
        return;
    }

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);
    formData.append('startHeaderRow', startHeaderRow);
    formData.append('endHeaderRow', endHeaderRow);
    formData.append('startDataRow', startDataRow);
    formData.append('endDataRow', endDataRow);
    formData.append('columnCount', columnCount);
    formData.append('templateName', templateName);

    fetch('http://localhost:10000/api/excel/upload', {
        method: 'POST',
        body: formData,
    })
        .then(async response => {
            if (!response.ok) {
                const errorText = await response.text(); // 获取错误信息
                throw new Error(errorText);
            }
            return response.text();
        })
        .then(data => {
            showPopup('上传成功: ' + data, 'success');
        })
        .catch(error => {
            showPopup('上传失败: ' + error.message, 'error');
        });
}

function downloadExcel() {
    const excelFileId = document.getElementById('excelFileIdDownload').value;

    if (!excelFileId) {
        showPopup('请输入文件 ID。', 'error');
        return;
    }

    fetch(`http://localhost:10000/api/excel/download?excelFileId=${encodeURIComponent(excelFileId)}`, {
        method: 'GET',
    })
        .then(async response => {
            if (!response.ok) {
                const errorText = await response.text(); // 获取错误信息
                throw new Error(errorText);
            }
            return response.blob();
        })
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `exported.xlsx`;
            document.body.appendChild(a);
            a.click();
            a.remove();
        })
        .catch(error => {
            showPopup('下载失败: ' + error.message, 'error');
        });
}
