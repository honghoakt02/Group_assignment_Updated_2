# 📈 HOSE Quant V4: Nền Tảng Kiểm Định Đầu Tư Định Lượng Cổ Phiếu HOSE

> **Ứng dụng Web App xây dựng trên nền tảng Streamlit nhằm kiểm định quy trình đầu tư định lượng, chọn lọc cổ phiếu tốt từ rổ HOSE, xây dựng danh mục tối ưu và đánh giá hiệu quả theo thời gian với cơ chế triệt tiêu hoàn toàn rò rỉ dữ liệu tương lai (No Look-Ahead Bias / Zero Future Leakage).**

---

## 🌟 1. Giới Thiệu & Mục Tiêu

Trong đầu tư định lượng (Quantitative Investing), vấn đề nguy hiểm nhất thường dẫn đến thất bại khi giao dịch thực tế là **rò rỉ dữ liệu tương lai (Look-Ahead Bias / Future Leakage)** và **quá khớp dữ liệu (Overfitting)**. 

Dự án này chuyển hóa toàn bộ pipeline nghiên cứu từ notebook **HOSE V4 Generic Nested Walk-Forward** thành một ứng dụng web app trực quan, linh hoạt và sẵn sàng triển khai trên **Streamlit Community Cloud** và **GitHub**.

### Các Trụ Cột Cốt Lõi Của Hệ Thống:
1. **Thiết kế phân chia dữ liệu nhân quả (Causal Data Split):**
   - **Core Train (60%):** Chỉ dùng để tối ưu hóa cặp ngưỡng kỹ thuật $(BUY, EXIT)$ chung cho toàn bộ universe thông qua K-Fold Cross Validation.
   - **Validation (20%):** Chạy Nested Walk-Forward để các họ mô hình phân bổ cạnh tranh sòng phẳng và khóa **Champion**.
   - **Final Test (20% / Năm cuối):** Hoàn toàn là kiểm định ngoài mẫu (Out-of-Sample - OOS). Tuyệt đối không được chỉnh sửa tham số sau khi nhìn kết quả Test.
2. **Tái chọn danh mục theo quý (Quarterly Nested Walk-Forward):**
   - Danh sách Top N cổ phiếu không bị khóa cứng cho toàn bộ giai đoạn kiểm định mà được tái lựa chọn vào đầu mỗi quý ($EffectiveDate$).
   - Tại ngày $EffectiveDate$, hệ thống chỉ được dùng dữ liệu giá và chỉ báo đến hết phiên trước ($AsOfDate = t-1$).
3. **Phân bổ danh mục đa mô hình (Portfolio Optimization):**
   - Hỗ trợ 5 họ phân bổ: **Equal Weight (EW)**, **MPT Ledoit-Wolf Shrinkage (MPT-LW)**, **Minimum Variance (MinVar-LW)**, **Hierarchical Risk Parity (HRP)**, và **Blend 50% MPT + 50% EW**.
   - Tái cân bằng tỷ trọng hàng tháng (Monthly Rebalance) dựa trên ma trận hiệp phương sai co rút (Ledoit-Wolf) dùng giá $\le t-1$.
4. **Cơ chế quản trị rủi ro & Đệm tiền mặt (Market Regime & Cash Buffer):**
   - Lọc trạng thái thị trường: Khi VN-Index dưới EMA200 hoặc độ dốc âm, tỷ trọng cổ phiếu chuyển về 0 và giữ tiền mặt (Cash).
   - Khi cổ phiếu kích hoạt tín hiệu bán ($EXIT$), tỷ trọng của mã đó không bị ép phân bổ lại sang mã khác mà được giữ ở dạng **Cash**.
5. **Cơ chế khớp lệnh thực tế:**
   - Tín hiệu sinh ra từ giá đóng cửa $Close(t-1)$ chỉ được phép khớp vào giá mở cửa phiên hôm sau $Open(t)$.
   - Tính toán đầy đủ chi phí giao dịch (phí môi giới $0.15\%$ và trượt giá $0.05\%$).

---

## 📁 2. Cấu Trúc Dự Án

Dự án bao gồm 3 file cốt lõi chuẩn mực để triển khai lên GitHub và Streamlit Cloud:

```text
├── app.py                                            # Ứng dụng Streamlit chính (UI + Quant Engine)
├── requirements.txt                                  # Danh sách thư viện Python cần cài đặt
├── README.md                                         # Tài liệu hướng dẫn sử dụng và triển khai
└── HOSE - 01.01.2021 - 31.12.2025 (97 ma).xlsx      # File dữ liệu OHLCV 97 mã cổ phiếu HOSE (mặc định)
```

- **`app.py`**: Chứa toàn bộ giao diện Streamlit tương tác, xử lý dữ liệu OHLCV, tính toán nhóm chỉ báo TA, tối ưu hóa ngưỡng Core Train, chạy mô phỏng sự kiện (Event-driven portfolio simulation), vẽ đồ thị tài sản Plotly trực quan và xuất file Excel kết quả.
- **`requirements.txt`**: Khai báo các thư viện phụ thuộc (`streamlit`, `pandas`, `numpy`, `scipy`, `scikit-learn`, `openpyxl`, `plotly`, `matplotlib`).
- **`README.md`**: Hướng dẫn chi tiết về cấu trúc, cách chạy local và quy trình deploy lên Streamlit Cloud.

---

## 🚀 3. Hướng Dẫn Cài Đặt & Chạy Ứng Dụng Ở Máy Cục Bộ (Local)

### Bước 1: Chuẩn bị môi trường Python
Khuyến nghị sử dụng Python phiên bản **3.10** hoặc **3.11**. Mở Terminal / PowerShell và di chuyển vào thư mục dự án:

```bash
cd "đường_dẫn_tới_thư_mục_chứa_code"
```

### Bước 2: Tạo và kích hoạt môi trường ảo (Khuyến nghị)
```bash
# Trên Windows:
python -m venv venv
venv\Scripts\activate

# Trên macOS/Linux:
python3 -m venv venv
source venv/bin/activate
```

### Bước 3: Cài đặt các thư viện cần thiết
```bash
pip install -r requirements.txt
```

### Bước 4: Khởi chạy ứng dụng Streamlit
```bash
streamlit run app.py
```
Ứng dụng sẽ tự động mở trên trình duyệt tại địa chỉ: `http://localhost:8501`.

---

## 🌐 4. Hướng Dẫn Đẩy Lên GitHub & Deploy Lên Streamlit Cloud

### Bước 1: Đẩy mã nguồn lên GitHub

1. Truy cập [GitHub](https://github.com) và tạo một Repository mới (ví dụ đặt tên: `hose-quant-backtest`).
2. Mở terminal tại thư mục dự án trên máy tính của bạn và chạy các lệnh:

```bash
git init
git add app.py requirements.txt README.md "HOSE - 01.01.2021 - 31.12.2025 (97 ma).xlsx"
git commit -m "Initial commit: Streamlit HOSE Quant Backtesting Web App V4"
git branch -M main
git remote add origin https://github.com/<tai-khoan-github-cua-ban>/hose-quant-backtest.git
git push -u origin main
```

*(Lưu ý: File dữ liệu Excel có kích thước khoảng ~9.2 MB, hoàn toàn nằm trong giới hạn 100 MB của GitHub).*

### Bước 2: Deploy ứng dụng lên Streamlit Community Cloud (Miễn phí)

1. Truy cập vào [share.streamlit.io](https://share.streamlit.io/) và đăng nhập bằng tài khoản GitHub của bạn.
2. Nhấn nút **"New app"** (hoặc **"Create app"**).
3. Điền thông tin cấu hình:
   - **Repository:** `<tai-khoan-github-cua-ban>/hose-quant-backtest`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Nhấn **"Deploy!"**.
5. Streamlit Cloud sẽ tự động cài đặt các dependencies từ `requirements.txt` và khởi chạy ứng dụng sau khoảng 1-2 phút. Bạn sẽ nhận được đường link web công khai dạng: `https://<ten-app>.streamlit.app` để chia sẻ cho mọi người.

---

## 📊 5. Cấu Trúc Các Tab Tính Năng Trên Web App

Ứng dụng được chia thành 7 Tab chuyên sâu:

| Tab | Tính Năng & Ý Nghĩa |
| :--- | :--- |
| **1. Dữ Liệu & Phân Kỳ** | Kiểm tra chất lượng dữ liệu 97 mã HOSE, tỷ lệ thiếu thanh khoản, bảng phân chia thời gian 3 tập Core Train / Validation / Final Test không chồng lấn. |
| **2. Ngưỡng TA Core Train** | Bảng kết quả Grid Search chọn cặp ngưỡng $(BUY, EXIT)$ tối ưu dựa trên điểm ổn định Sharpe CV trên Core Train. |
| **3. Top N Walk-Forward** | Lịch sử tái chọn Top N cổ phiếu theo từng quý tại các phiên $t-1$, kèm bảng xếp hạng đa yếu tố chi tiết cho từng kỳ. |
| **4. Khóa Champion (Validation)** | Cuộc đua hiệu quả giữa 5 mô hình phân bổ danh mục (EW, MPT, MinVar, HRP, Blend) trên Validation và khóa Champion chiến thắng. |
| **5. Kiểm Định OOS Final Test** | Bảng chỉ số hiệu quả ngoài mẫu (Return, CAGR, Sharpe, Sortino, Max Drawdown, Calmar, Turnover, Exposure) và đồ thị NAV so sánh với VN-Index. |
| **6. Kiểm Toán Rò Rỉ Tương Lai** | Ma trận kiểm toán 7 bước chứng minh mô hình hoàn toàn tuân thủ tính nhân quả và không có Look-ahead bias. |
| **7. Xuất Báo Cáo** | Nút tải xuống file Excel đa sheet (`HOSE_V4_RESULTS.xlsx`) chứa đầy đủ các bảng dữ liệu để trình bày báo cáo hoặc nghiên cứu tiếp theo. |

---

## 🛡️ 6. Bảng Kiểm Toán Rò Rỉ Thông Tin Tương Lai (Leakage Audit)

| Quy Trình / Giai Đoạn | Mốc Dữ Liệu Tối Đa Sử Dụng | Cho Phép Retune? | Nguyên Tắc Chống Rò Rỉ |
| :--- | :---: | :---: | :--- |
| **1. TA BUY / EXIT Selection** | `Core Train End` | ❌ Không | Ngưỡng kỹ thuật chỉ được tối ưu hóa trên Core Train qua K-Fold CV. |
| **2. Model Family Selection** | `Validation End` | ❌ Không | Họ mô hình phân bổ danh mục (EW, MPT, MinVar...) được khóa tại Validation. |
| **3. Quarterly Top N Reselection** | $t - 1$ | ❌ Không | Top N tại phiên $t$ chỉ dùng chỉ báo và giá đến hết phiên $t-1$. |
| **4. Monthly Weight Rebalancing** | $t - 1$ | ❌ Không | Ma trận hiệp phương sai Ledoit-Wolf chỉ tính bằng dữ liệu giá $\le t-1$. |
| **5. Execution Timing** | $Close(t-1) \to Open(t)$ | ❌ Không | Tín hiệu sinh tại giá đóng cửa $t-1$ chỉ được khớp vào giá mở cửa $t$. |
| **6. Market Regime / Cash** | $Close(t-1) \to Open(t)$ | ❌ Không | Bộ lọc xu hướng thị trường tính tại $t-1$. Khi thị trường giảm, vốn chuyển sang Cash. |
| **7. Final Test Evaluation** | `Final Test End` | ❌ Không | Final Test là tập OOS thuần túy; hoàn toàn không thay đổi tham số sau khi chạy. |

---

## 👨‍💻 Tác Quyền & Liên Hệ
- **Quy trình định lượng:** HOSE V4 Nested Walk-Forward Quantitative Pipeline.
- **Nền tảng:** Python, Streamlit, Plotly, Scikit-learn, Scipy, Pandas.
