// QCR v4.0 前端脚本
console.log('QCR v4.0 已加载');

// 文件验证
document.addEventListener('DOMContentLoaded', function() {
    const fileInputs = document.querySelectorAll('input[type="file"]');
    fileInputs.forEach(input => {
        input.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (file) {
                const maxSize = 100 * 1024 * 1024; // 100MB
                if (file.size > maxSize) {
                    alert('文件太大！最大100MB');
                    e.target.value = '';
                }
            }
        });
    });
});

