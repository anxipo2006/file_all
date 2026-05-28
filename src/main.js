import './styles.css';
const API_BASE = import.meta.env.VITE_API_BASE_URL || 'http://localhost:8000';
const CATS={
 '🧠 Master Tool':[['master_tool','Master auto pipeline']],
 '🗜️ Nén & chia AI':[['compress','Nén file'],['compress_split_pdf','Nén + chia PDF'],['split_pdf','Chia PDF theo MB']],
 '🌐 Dịch file':[['translate_docx','Dịch Word'],['translate_xlsx','Dịch Excel'],['translate_pptx','Dịch PowerPoint']],
 '🔄 Chuyển đổi':[['images_to_pdf','Ảnh → PDF'],['pdf_to_docx','PDF → Word'],['docx_to_pdf','Word → PDF'],['pptx_to_pdf','PPT → PDF']],
 '📄 PDF tools':[['merge_pdfs','Gộp PDF'],['extract_pages','Lấy trang'],['delete_pages','Xóa trang'],['rotate_pdf','Xoay PDF'],['pdf_to_images','PDF → ảnh ZIP'],['extract_text_pdf','Trích text'],['pdf_metadata','Metadata']],
 '🖼️ Image tools':[['convert_image','Đổi format ảnh'],['resize_image','Resize ảnh'],['compress_image','Nén ảnh']],
 '📦 Archive/Split':[['zip_files','ZIP nhiều file'],['split_binary','Chia file bất kỳ']]
};
const $=s=>document.querySelector(s); const cat=$('#category'), action=$('#action'), form=$('#tool-form'), file=$('#file-input'), fname=$('#file-name'), status=$('#status'), metrics=$('#metrics'), dl=$('#download-link'), btn=$('#submit-btn');
Object.keys(CATS).forEach(k=>cat.add(new Option(k,k))); function fill(){action.innerHTML=''; CATS[cat.value].forEach(([v,t])=>action.add(new Option(t,v))); $('#goal-wrap').style.display=action.value==='master_tool'?'grid':'none';} cat.onchange=fill; action.onchange=fill; fill();
file.onchange=()=>{const fs=[...file.files]; fname.textContent=fs.length?fs.map(f=>`${f.name} • ${fmt(f.size)}`).join(' | '):'Chưa chọn file'};
form.onsubmit=async e=>{e.preventDefault(); if(!file.files.length)return; btn.disabled=true; btn.textContent='⏳ Đang xử lý...'; status.textContent='Đang upload/xử lý. File lớn có thể mất lâu.'; metrics.innerHTML=''; dl.classList.add('hidden');
 const data=new FormData(); data.append('action',action.value); data.append('goal',$('#goal').value); data.append('target_mb',$('#target-mb').value); data.append('max_part_mb',$('#part-mb').value); data.append('target_lang',$('#target-lang').value); data.append('page_range',$('#page-range').value); data.append('rotate_degrees',$('#rotate').value); data.append('output_format',$('#output-format').value); data.append('max_width',$('#max-width').value);
 [...file.files].forEach((f,i)=>data.append(i===0?'file':'files',f)); if(file.files.length===1)data.append('files',file.files[0]);
 try{const res=await fetch(`${API_BASE}/tool`,{method:'POST',body:data}); if(!res.ok){const er=await res.json().catch(()=>({error:'Lỗi không xác định'})); throw new Error(er.error)} const blob=await res.blob(); const url=URL.createObjectURL(blob); const name=filename(res.headers.get('content-disposition'))||'result.bin'; dl.href=url; dl.download=name; dl.textContent=`⬇️ Tải ${name}`; dl.classList.remove('hidden'); metrics.innerHTML=`<div><strong>Action</strong><span>${res.headers.get('x-tool-action')||action.value}</span></div><div><strong>Plan</strong><span>${res.headers.get('x-master-plan')||'-'}</span></div><div><strong>Trước</strong><span>${res.headers.get('x-input-size-text')||'-'}</span></div><div><strong>Sau</strong><span>${res.headers.get('x-output-size-text')||fmt(blob.size)}</span></div><div><strong>Giảm</strong><span>${res.headers.get('x-saved-percent')||'0'}%</span></div>`; status.textContent='Hoàn tất.';}catch(err){status.textContent=`Lỗi: ${err.message}`;}finally{btn.disabled=false; btn.textContent='🚀 Chạy tool';}}
function fmt(b){if(b<1024)return`${b} B`; if(b<1048576)return`${(b/1024).toFixed(1)} KB`; return`${(b/1048576).toFixed(2)} MB`}
function filename(d){return /filename="?([^";]+)"?/i.exec(d||'')?.[1]}
