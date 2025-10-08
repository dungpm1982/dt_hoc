# python.py

import streamlit as st
import pandas as pd
from google import genai
from google.genai.errors import APIError
import os # Import thư viện os để đọc biến môi trường

# --- Cấu hình Trang Streamlit ---
st.set_page_config(
    page_title="App Phân Tích Báo Cáo Tài Chính",
    layout="wide"
)

st.title("Ứng dụng Phân Tích Báo Cáo Tài Chính 📊")

# =========================================================================
# --- KHỐI BỔ SUNG: KHỞI TẠO GEMINI CLIENT & QUẢN LÝ TRẠNG THÁI CHAT ---
# =========================================================================

# Lấy khóa API từ Streamlit Secrets (Ưu tiên) hoặc biến môi trường
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY")

if not GEMINI_API_KEY:
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

CLIENT = None
if GEMINI_API_KEY:
    try:
        # Khởi tạo client chung cho cả ứng dụng
        CLIENT = genai.Client(api_key=GEMINI_API_KEY)
        MODEL = "gemini-2.5-flash"
    except Exception as e:
        st.error(f"Lỗi khởi tạo Gemini Client: {e}")
        
    # Khởi tạo trạng thái chat chỉ khi client được tạo thành công
    if CLIENT and "messages" not in st.session_state:
        # Lịch sử chat hiển thị
        st.session_state.messages = [{"role": "model", "content": "Chào bạn! Tôi là trợ lý AI. Bạn có câu hỏi nào về phân tích tài chính hoặc cần hỗ trợ gì không?"}]

    if CLIENT and "chat_session" not in st.session_state:
        # Phiên chat để duy trì bối cảnh hội thoại
        st.session_state.chat_session = CLIENT.chats.create(model=MODEL)
        
# =========================================================================
# --- END KHỐI BỔ SUNG ---
# =========================================================================

# --- Hàm tính toán chính (Sử dụng Caching để Tối ưu hiệu suất) ---
@st.cache_data
def process_financial_data(df):
    """Thực hiện các phép tính Tăng trưởng và Tỷ trọng."""
    
    # Đảm bảo các giá trị là số để tính toán
    numeric_cols = ['Năm trước', 'Năm sau']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # 1. Tính Tốc độ Tăng trưởng
    # Dùng .replace(0, 1e-9) cho Series Pandas để tránh lỗi chia cho 0
    df['Tốc độ tăng trưởng (%)'] = (
        (df['Năm sau'] - df['Năm trước']) / df['Năm trước'].replace(0, 1e-9)
    ) * 100

    # 2. Tính Tỷ trọng theo Tổng Tài sản
    # Lọc chỉ tiêu "TỔNG CỘNG TÀI SẢN"
    tong_tai_san_row = df[df['Chỉ tiêu'].str.contains('TỔNG CỘNG TÀI SẢN', case=False, na=False)]
    
    if tong_tai_san_row.empty:
        raise ValueError("Không tìm thấy chỉ tiêu 'TỔNG CỘNG TÀI SẢN'.")

    tong_tai_san_N_1 = tong_tai_san_row['Năm trước'].iloc[0]
    tong_tai_san_N = tong_tai_san_row['Năm sau'].iloc[0]

    # ******************************* PHẦN SỬA LỖI BẮT ĐẦU *******************************
    # Lỗi xảy ra khi dùng .replace() trên giá trị đơn lẻ (numpy.int64).
    # Sử dụng điều kiện ternary để xử lý giá trị 0 thủ công cho mẫu số.
    
    divisor_N_1 = tong_tai_san_N_1 if tong_tai_san_N_1 != 0 else 1e-9
    divisor_N = tong_tai_san_N if tong_tai_san_N != 0 else 1e-9

    # Tính tỷ trọng với mẫu số đã được xử lý
    df['Tỷ trọng Năm trước (%)'] = (df['Năm trước'] / divisor_N_1) * 100
    df['Tỷ trọng Năm sau (%)'] = (df['Năm sau'] / divisor_N) * 100
    # ******************************* PHẦN SỬA LỖI KẾT THÚC *******************************
    
    return df

# --- Hàm gọi API Gemini ---
def get_ai_analysis(data_for_ai, api_key):
    """Gửi dữ liệu phân tích đến Gemini API và nhận nhận xét."""
    try:
        # Cần tạo client riêng ở đây vì hàm này được gọi độc lập với chat
        client_analysis = genai.Client(api_key=api_key)
        model_name = 'gemini-2.5-flash' 

        prompt = f"""
        Bạn là một chuyên gia phân tích tài chính chuyên nghiệp. Dựa trên các chỉ số tài chính sau, hãy đưa ra một nhận xét khách quan, ngắn gọn (khoảng 3-4 đoạn) về tình hình tài chính của doanh nghiệp. Đánh giá tập trung vào tốc độ tăng trưởng, thay đổi cơ cấu tài sản và khả năng thanh toán hiện hành.
        
        Dữ liệu thô và chỉ số:
        {data_for_ai}
        """

        response = client_analysis.models.generate_content(
            model=model_name,
            contents=prompt
        )
        return response.text

    except APIError as e:
        return f"Lỗi gọi Gemini API: Vui lòng kiểm tra Khóa API hoặc giới hạn sử dụng. Chi tiết lỗi: {e}"
    except KeyError:
        return "Lỗi: Không tìm thấy Khóa API 'GEMINI_API_KEY'. Vui lòng kiểm tra cấu hình Secrets trên Streamlit Cloud."
    except Exception as e:
        return f"Đã xảy ra lỗi không xác định: {e}"


# --- Chức năng 1: Tải File ---
uploaded_file = st.file_uploader(
    "1. Tải file Excel Báo cáo Tài chính (Chỉ tiêu | Năm trước | Năm sau)",
    type=['xlsx', 'xls']
)

# Khởi tạo giá trị mặc định cho chỉ số thanh toán để tránh lỗi tham chiếu
thanh_toan_hien_hanh_N = "N/A"
thanh_toan_hien_hanh_N_1 = "N/A"

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        # Tiền xử lý: Đảm bảo chỉ có 3 cột quan trọng
        df_raw.columns = ['Chỉ tiêu', 'Năm trước', 'Năm sau']
        
        # Xử lý dữ liệu
        df_processed = process_financial_data(df_raw.copy())

        if df_processed is not None:
            
            # --- Chức năng 2 & 3: Hiển thị Kết quả ---
            st.subheader("2. Tốc độ Tăng trưởng & 3. Tỷ trọng Cơ cấu Tài sản")
            st.dataframe(df_processed.style.format({
                'Năm trước': '{:,.0f}',
                'Năm sau': '{:,.0f}',
                'Tốc độ tăng trưởng (%)': '{:.2f}%',
                'Tỷ trọng Năm trước (%)': '{:.2f}%',
                'Tỷ trọng Năm sau (%)': '{:.2f}%'
            }), use_container_width=True)
            
            # --- Chức năng 4: Tính Chỉ số Tài chính ---
            st.subheader("4. Các Chỉ số Tài chính Cơ bản")
            
            try:
                # Lấy Tài sản ngắn hạn
                tsnh_n = df_processed[df_processed['Chỉ tiêu'].str.contains('TÀI SẢN NGẮN HẠN', case=False, na=False)]['Năm sau'].iloc[0]
                tsnh_n_1 = df_processed[df_processed['Chỉ tiêu'].str.contains('TÀI SẢN NGẮN HẠN', case=False, na=False)]['Năm trước'].iloc[0]

                # Lấy Nợ ngắn hạn
                no_ngan_han_N = df_processed[df_processed['Chỉ tiêu'].str.contains('NỢ NGẮN HẠN', case=False, na=False)]['Năm sau'].iloc[0]  
                no_ngan_han_N_1 = df_processed[df_processed['Chỉ tiêu'].str.contains('NỢ NGẮN HẠN', case=False, na=False)]['Năm trước'].iloc[0]
                
                # Tránh chia cho 0 khi tính toán
                if no_ngan_han_N != 0:
                    thanh_toan_hien_hanh_N = tsnh_n / no_ngan_han_N
                if no_ngan_han_N_1 != 0:
                    thanh_toan_hien_hanh_N_1 = tsnh_n_1 / no_ngan_han_N_1


                col1, col2 = st.columns(2)
                with col1:
                    st.metric(
                        label="Chỉ số Thanh toán Hiện hành (Năm trước)",
                        value=f"{thanh_toan_hien_hanh_N_1:.2f} lần" if isinstance(thanh_toan_hien_hanh_N_1, float) else "N/A"
                    )
                with col2:
                    current_ratio_N_val = f"{thanh_toan_hien_hanh_N:.2f} lần" if isinstance(thanh_toan_hien_hanh_N, float) else "N/A"
                    delta_val = thanh_toan_hien_hanh_N - thanh_toan_hien_hanh_N_1 if isinstance(thanh_toan_hien_hanh_N, float) and isinstance(thanh_toan_hien_hanh_N_1, float) else None
                    
                    st.metric(
                        label="Chỉ số Thanh toán Hiện hành (Năm sau)",
                        value=current_ratio_N_val,
                        delta=f"{delta_val:.2f}" if delta_val is not None else None
                    )
                    
            except IndexError:
                st.warning("Thiếu chỉ tiêu 'TÀI SẢN NGẮN HẠN' hoặc 'NỢ NGẮN HẠN' để tính chỉ số.")
                thanh_toan_hien_hanh_N = "N/A" 
                thanh_toan_hien_hanh_N_1 = "N/A"
            except Exception as e:
                 st.warning(f"Lỗi khi tính chỉ số tài chính: {e}")
                 thanh_toan_hien_hanh_N = "N/A" 
                 thanh_toan_hien_hanh_N_1 = "N/A"
            
            # --- Chức năng 5: Nhận xét AI ---
            st.subheader("5. Nhận xét Tình hình Tài chính (AI)")
            
            # Chuẩn bị dữ liệu để gửi cho AI
            # Cần xử lý trường hợp N/A cho chỉ số thanh toán trước khi gửi
            thanh_toan_N_val = f"{thanh_toan_hien_hanh_N:.2f}" if isinstance(thanh_toan_hien_hanh_N, float) else "Không xác định"
            thanh_toan_N_1_val = f"{thanh_toan_hien_hanh_N_1:.2f}" if isinstance(thanh_toan_hien_hanh_N_1, float) else "Không xác định"

            try:
                tsnh_tang_truong = df_processed[df_processed['Chỉ tiêu'].str.contains('TÀI SẢN NGẮN HẠN', case=False, na=False)]['Tốc độ tăng trưởng (%)'].iloc[0]
                tsnh_tang_truong_val = f"{tsnh_tang_truong:.2f}%"
            except IndexError:
                tsnh_tang_truong_val = "Không xác định"

            data_for_ai = pd.DataFrame({
                'Chỉ tiêu': [
                    'Toàn bộ Bảng phân tích (dữ liệu thô)', 
                    'Tăng trưởng Tài sản ngắn hạn (%)', 
                    'Thanh toán hiện hành (N-1)', 
                    'Thanh toán hiện hành (N)'
                ],
                'Giá trị': [
                    df_processed.to_markdown(index=False),
                    tsnh_tang_truong_val,
                    thanh_toan_N_1_val,
                    thanh_toan_N_val
                ]
            }).to_markdown(index=False) 

            if st.button("Yêu cầu AI Phân tích"):
                if GEMINI_API_KEY:
                    with st.spinner('Đang gửi dữ liệu và chờ Gemini phân tích...'):
                        # Gọi hàm phân tích chuyên sâu
                        ai_result = get_ai_analysis(data_for_ai, GEMINI_API_KEY)
                    st.markdown("**Kết quả Phân tích từ Gemini AI:**")
                    st.info(ai_result)
                else:
                    st.error("Lỗi: Không tìm thấy Khóa API. Vui lòng cấu hình Khóa 'GEMINI_API_KEY' trong Streamlit Secrets.")

    except ValueError as ve:
        st.error(f"Lỗi cấu trúc dữ liệu: {ve}")
    except Exception as e:
        st.error(f"Có lỗi xảy ra khi đọc hoặc xử lý file: {e}. Vui lòng kiểm tra định dạng file.")

else:
    st.info("Vui lòng tải lên file Excel để bắt đầu phân tích.")

# =========================================================================
# --- KHỐI BỔ SUNG: KHUNG CHAT HỎI ĐÁP VỚI GEMINI (Độc lập) ---
# =========================================================================

st.markdown("---") # Đường kẻ phân cách

if CLIENT:
    st.subheader("6. Trợ lý Hỏi đáp Gemini")
    st.markdown("Bạn có thể hỏi trực tiếp trợ lý AI về bất kỳ vấn đề nào hoặc yêu cầu làm rõ các chỉ số tài chính.")
    
    # 1. Hiển thị lịch sử chat
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    # 2. Xử lý input
    user_prompt = st.chat_input("Hỏi Gemini...")

    if user_prompt:
        # Thêm tin nhắn user vào lịch sử
        st.session_state.messages.append({"role": "user", "content": user_prompt})

        # Hiển thị tin nhắn user ngay lập tức
        with st.chat_message("user"):
            st.markdown(user_prompt)

        # Lấy phản hồi từ Gemini
        with st.chat_message("model"):
            with st.spinner("Gemini đang trả lời..."):
                try:
                    # Sử dụng chat session đã được lưu trong session_state
                    response = st.session_state.chat_session.send_message(user_prompt)
                    st.markdown(response.text)
                    # Thêm phản hồi của Gemini vào lịch sử
                    st.session_state.messages.append({"role": "model", "content": response.text})
                except APIError as e:
                    error_msg = f"Lỗi API: Không thể nhận phản hồi từ Gemini. Vui lòng kiểm tra Khóa API hoặc giới hạn sử dụng. Chi tiết lỗi: {e}"
                    st.error(error_msg)
                    st.session_state.messages.append({"role": "model", "content": error_msg})
                except Exception as e:
                    error_msg = f"Đã xảy ra lỗi không xác định: {e}"
                    st.error(error_msg)
                    st.session_state.messages.append({"role": "model", "content": error_msg})
else:
    # Thông báo nếu API Key không được cấu hình
    st.warning("Tính năng Trợ lý Hỏi đáp Gemini không hoạt động do thiếu Khóa API. Vui lòng cấu hình GEMINI_API_KEY.")
# =========================================================================
# --- END KHỐI BỔ SUNG ---
# =========================================================================
