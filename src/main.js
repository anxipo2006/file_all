import './styles.css';

const API_BASE = import.meta.env.VITE_API_BASE_URL || 'http://localhost:8000';

const form = document.querySelector('#file-form');
const fileInput = document.querySelector('#file-input');
const fileName = document.querySelector('#file-name');
const statusEl = document.querySelector('#status');
const metricsEl = document.querySelector('#metrics');
const downloadLink = document.querySelector('#download-link');
const submitBtn = document.querySelector('#submit-btn');

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  fileName.textContent = file ? `${file.name} • ${formatBytes(file.size)}` : 'Chưa chọn file';
});

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  const file = fileInput.files?.[0];
  if (!file) return;

  submitBtn.disabled = true;
  submitBtn.textContent = '⏳ Đang xử lý...';
  statusEl.textContent = 'Đang upload và xử lý file. File lớn có thể mất vài phút.';
  metricsEl.innerHTML = '';
  downloadLink.classList.add('hidden');

  const data = new FormData();
  data.append('file', file);
  data.append('target_mb', document.querySelector('#target-mb').value);
  data.append('max_part_mb', document.querySelector('#part-mb').value);
  data.append('split_for_ai', document.querySelector('#split-ai').checked);
  data.append('pdf_mode', document.querySelector('#pdf-mode').value);

  try {
    const response = await fetch(`${API_BASE}/process`, {
      method: 'POST',
      body: data,
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({ error: 'Lỗi không xác định' }));
      throw new Error(error.error || 'Xử lý thất bại');
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const disposition = response.headers.get('content-disposition') || '';
    const filename = extractFilename(disposition) || `processed_${file.name}`;

    downloadLink.href = url;
    downloadLink.download = filename;
    downloadLink.classList.remove('hidden');
    downloadLink.textContent = `⬇️ Tải ${filename}`;

    metricsEl.innerHTML = `
      <div><strong>Trước</strong><span>${response.headers.get('x-input-size-text') || formatBytes(file.size)}</span></div>
      <div><strong>Sau</strong><span>${response.headers.get('x-output-size-text') || formatBytes(blob.size)}</span></div>
      <div><strong>Giảm</strong><span>${response.headers.get('x-saved-percent') || '0'}%</span></div>
    `;
    statusEl.textContent = 'Hoàn tất. Bạn có thể tải file kết quả.';
  } catch (error) {
    statusEl.textContent = `Lỗi: ${error.message}`;
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = '🚀 Xử lý file';
  }
});

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}

function extractFilename(disposition) {
  const match = disposition.match(/filename="?([^";]+)"?/i);
  return match?.[1];
}
