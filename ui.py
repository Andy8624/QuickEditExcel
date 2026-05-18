import io
import zipfile

import streamlit as st

from excel_logic import extract_class_from_filename, process_excel

# --- BỘ CSS MỚI: HIỆN ĐẠI & CHUYÊN NGHIỆP HƠN ---
BASE_CSS = """
<style>
/* Đổi màu nền toàn trang sang xám nhạt */
.stApp {
    background-color: #f3f6f9;
}

/* Biến container thành một khối Card trắng, có đổ bóng */
.block-container {
    padding: 2.5rem 2rem 3rem 2rem !important;
    max-width: 760px;
    background-color: #ffffff;
    border-radius: 16px;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04);
    margin-top: 3rem;
    margin-bottom: 3rem;
}

/* Ẩn header mặc định của Streamlit */
[data-testid="stHeaderActionElements"] {
    display: none !important;
}

/* Phần Tiêu đề (Hero) */
.hero {
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 24px;
    padding-bottom: 16px;
    border-bottom: 1px solid #e2e8f0;
}
.hero-icon {
    width: 56px;
    height: 56px;
    border-radius: 14px;
    display: grid;
    place-items: center;
    background: linear-gradient(135deg, #3b82f6, #4f46e5);
    color: #ffffff;
    font-size: 28px;
    box-shadow: 0 4px 10px rgba(59, 130, 246, 0.3);
}
.hero-title {
    margin: 0;
    color: #0f172a;
    font-weight: 800;
    font-size: clamp(24px, 4vw, 36px);
    line-height: 1.2;
}

/* Tùy chỉnh vùng Upload File (Dropzone) */
[data-testid="stFileUploader"] {
    margin: 1rem 0;
    width: 100% !important;
}
[data-testid="stFileUploader"] label[data-testid="stWidgetLabel"],
[data-testid="stFileUploaderFileList"] {
    display: none !important;
}
[data-testid="stFileUploaderDropzone"] {
    position: relative;
    width: 100% !important;
    min-height: 100px !important; /* Mở rộng vùng kéo thả */
    border: 2px dashed #cbd5e1 !important;
    border-radius: 12px !important;
    background: #f8fafc !important;
    padding: 0 !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    overflow: hidden;
    cursor: pointer;
    transition: all 0.2s ease-in-out;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #3b82f6 !important;
    background: #eff6ff !important;
}
[data-testid="stFileUploaderDropzone"] > div,
[data-testid="stFileUploaderDropzone"] button,
[data-testid="stFileUploaderDropzone"] [data-testid="stFileChips"],
[data-testid="stFileUploaderDropzone"] [data-testid="stFileChip"] {
    display: none !important;
}
[data-testid="stFileUploaderDropzone"] input[data-testid="stFileUploaderDropzoneInput"] {
    display: block !important;
    opacity: 0 !important;
    position: absolute !important;
    inset: 0 !important;
    width: 100% !important;
    height: 100% !important;
    cursor: pointer !important;
}
[data-testid="stFileUploaderDropzone"]::before {
    content: "📂 Kéo thả file Excel vào đây hoặc Click để chọn";
    pointer-events: none;
    color: #64748b;
    font-size: 15px;
    font-weight: 600;
    text-align: center;
}
[data-testid="stFileUploaderDropzone"]:hover::before {
    color: #2563eb;
}

/* Danh sách file đã tải lên */
.fl-wrap {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 12px;
    margin-top: 8px;
    max-height: 280px;
    overflow-y: auto;
}
.fl-item {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 10px 8px;
    border-bottom: 1px solid #f1f5f9;
    font-size: 14px;
}
.fl-item:last-child {
    border-bottom: none;
}
.fl-num {
    font-weight: 600;
    color: #94a3b8;
    min-width: 20px;
}
.fl-tag {
    background: #e0e7ff;
    color: #4f46e5;
    border-radius: 6px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 700;
    white-space: nowrap;
}
.fl-tag-err {
    background: #fee2e2;
    color: #dc2626;
    border-radius: 6px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 700;
    white-space: nowrap;
}
.fl-name {
    flex: 1;
    color: #334155;
    font-weight: 500;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.fl-total {
    font-size: 13px;
    color: #64748b;
    margin-top: 12px;
    padding-top: 12px;
    border-top: 1px dashed #e2e8f0;
    display: flex;
    justify-content: space-between;
}

/* Chỉnh nút bấm Streamlit */
.stButton, div[data-testid="stButton"] {
    width: 100%;
}
.stButton > button[kind="primary"],
div.stButton > button[kind="primary"] {
    width: 100%;
    height: 52px;
    font-size: 16px;
    font-weight: 600;
    border-radius: 12px;
    border: none;
    background: #0f172a; /* Màu tối hiện đại thay vì xanh mòng két */
    color: #ffffff;
    transition: all 0.2s;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}
div.stButton > button[kind="primary"]:hover {
    background: #334155;
    transform: translateY(-1px);
    box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
}

/* Responsive cho Mobile */
@media (max-width: 600px) {
    .block-container {
        padding: 1.5rem 1rem !important;
        margin-top: 1rem;
        border-radius: 0;
    }
    .hero-title {
        font-size: 22px;
    }
    [data-testid="stFileUploaderDropzone"]::before {
        font-size: 14px;
        content: "📂 Chạm để chọn file Excel";
    }
}
</style>
"""


def _reset_after_zip_download():
    """Xóa trạng thái uploader sau khi tải ZIP để bắt đầu lượt xử lý mới."""
    st.session_state["uploader_key"] = st.session_state.get("uploader_key", 0) + 1

def _render_uploaded_file_list(uploaded_files):
    items_html = ""
    invalid_count = 0

    for idx, file_obj in enumerate(uploaded_files, start=1):
        class_name, class_error = extract_class_from_filename(file_obj.name)
        if class_name:
            tag = f'<span class="fl-tag">{class_name}</span>'
        else:
            tag = '<span class="fl-tag-err">Lỗi tên</span>'
            invalid_count += 1

        items_html += (
            '<div class="fl-item">'
            f'<span class="fl-num">{idx}.</span>'
            f"{tag}"
            f'<span class="fl-name" title="{file_obj.name}">{file_obj.name}</span>'
            '</div>'
        )

    summary = f'<div class="fl-total"><span>Tổng: <b>{len(uploaded_files)}</b> file</span>'
    if invalid_count:
        summary += f'<span style="color:#dc2626; font-weight:600;">{invalid_count} lỗi tên lớp</span>'
    summary += '</div>'

    st.markdown(f'<div class="fl-wrap">{items_html}{summary}</div>', unsafe_allow_html=True)


def render_app():
    st.set_page_config(page_title="Chỉnh sửa học bạ Excel", page_icon="📚", layout="centered")
    st.markdown(BASE_CSS, unsafe_allow_html=True)

    if "uploader_key" not in st.session_state:
        st.session_state["uploader_key"] = 0

    st.markdown(
        """
        <div class="hero">
            <div class="hero-icon">📚</div>
            <h1 class="hero-title">Chỉnh sửa học bạ</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Thêm file Excel",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key=f"uploader_{st.session_state['uploader_key']}",
    )

    if uploaded_files:
        _render_uploaded_file_list(uploaded_files)
        st.write("")

    if st.button("🚀 Bắt đầu xử lý", type="primary", use_container_width=True, disabled=not uploaded_files):
        if not uploaded_files:
            st.warning("⚠️ Vui lòng tải lên ít nhất 1 file Excel.")
            return

        with st.spinner("Đang xử lý dữ liệu..."):
            processed_files = []
            errors = []

            for file_obj in uploaded_files:
                class_name, class_error = extract_class_from_filename(file_obj.name)
                if class_error:
                    errors.append(f"❌ **{file_obj.name}**: {class_error}")
                    continue

                out_stream, err_msg = process_excel(file_obj, class_name, file_obj.name)
                if err_msg:
                    errors.append(err_msg)
                else:
                    processed_files.append((file_obj.name, out_stream))

        st.divider()

        total = len(uploaded_files)
        success_count = len(processed_files)
        fail_count = len(errors)

        col_total, col_success, col_fail = st.columns(3)
        col_total.metric("📁 Tổng file", total)
        col_success.metric("✅ Thành công", success_count)
        col_fail.metric("❌ Thất bại", fail_count)

        if errors:
            with st.expander(f"Xem chi tiết {fail_count} file lỗi", expanded=fail_count > 0 and success_count == 0):
                for err in errors:
                    st.error(err)

        if not processed_files:
            return

        st.success(f"🎉 Đã xử lý thành công {success_count}/{total} file.")

        if len(processed_files) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for file_name, file_stream in processed_files:
                    zip_file.writestr(file_name, file_stream.getvalue())

            st.download_button(
                label=f"📦 Tải {success_count} file đã xử lý (.zip)",
                data=zip_buffer.getvalue(),
                file_name="Hoc_ba_da_xu_ly.zip",
                mime="application/zip",
                use_container_width=True,
                on_click=_reset_after_zip_download,
            )
        else:
            file_name, file_stream = processed_files[0]
            st.download_button(
                label="💾 Tải file đã xử lý (.xlsx)",
                data=file_stream.getvalue(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

if __name__ == "__main__":
    render_app()