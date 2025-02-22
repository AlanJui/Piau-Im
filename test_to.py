import cv2
import numpy as np

# 顯示處理後的圖片
# from PIL import Image

# 讀取圖片
# image = cv2.imread("/mnt/data/A_traditional_Chinese_painting-style_illustration_.png")
image = cv2.imread("A_traditional_Chinese_painting-style_illustration.png")

# 轉換為灰階圖像
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# 應用閾值處理來突出題字與印章
_, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

# 尋找題字與印章的輪廓
contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

# 創建遮罩
mask = np.zeros_like(image)

# 在遮罩上填充找到的區域
for contour in contours:
    x, y, w, h = cv2.boundingRect(contour)
    if w > 50 and h > 20:  # 過濾小區域，避免誤刪細節
        cv2.rectangle(mask, (x, y), (x + w, y + h), (255, 255, 255), -1)

# 使用 inpainting 修復去除的區域
result = cv2.inpaint(image, cv2.cvtColor(mask, cv2.COLOR_BGR2GRAY), 5, cv2.INPAINT_TELEA)

# 保存處理後的圖片
# output_path = "/mnt/data/cleaned_traditional_chinese_painting.png"
output_path = "cleaned_traditional_chinese_painting.png"
cv2.imwrite(output_path, result)

# cleaned_image = Image.open(output_path)
# cleaned_image.show()
cv2.imshow("Cleaned Image", result)
cv2.waitKey(0)
cv2.destroyAllWindows()