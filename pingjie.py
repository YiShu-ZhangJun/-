import fitz  # PyMuPDF
from PIL import Image
import os
import sys

def pdf_to_image(pdf_path, output_image_path, page_number, dpi=300):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_number)  # 0-based index
    zoom = dpi / 72  # 72是默认的DPI
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    pix.save(output_image_path)

def merge_images(image1_path, image2_path, output_path):
    # 加载两张图片
    image1 = Image.open(image1_path)
    image2 = Image.open(image2_path) if image2_path else None

    # 创建一张A4纸大小的空白图像 (2480x3508像素，300dpi)
    a4_width, a4_height = 2480, 3508
    a4_image = Image.new('RGB', (a4_width, a4_height), 'white')

    # 调整图片大小
    target_width = a4_width
    target_height = int(a4_height / 2)

    image1_resized = image1.resize((target_width, target_height))
    a4_image.paste(image1_resized, (0, 0))

    if image2:
        image2_resized = image2.resize((target_width, target_height))
        a4_image.paste(image2_resized, (0, target_height))

    # 保存拼接后的图像
    a4_image.save(output_path)

def main():
    # 接收命令行参数
    original_invoices_path = sys.argv[1]
    merge_output_path = sys.argv[2]

    # 确保合并输出目录存在
    os.makedirs(merge_output_path, exist_ok=True)

    # 获取原始发票目录中的所有PDF文件
    pdf_files = [os.path.join(original_invoices_path, f) for f in os.listdir(original_invoices_path) if f.endswith('.pdf')]

    # 将每个PDF文件转换为高质量图像
    images = []
    for i, pdf_file in enumerate(pdf_files):
        image_path = os.path.join(original_invoices_path, f"page_{i + 1}.png")
        pdf_to_image(pdf_file, image_path, 0, dpi=300)
        images.append(image_path)

    # 两两合并图像
    for i in range(0, len(images), 2):
        image1 = images[i]
        image2 = images[i + 1] if i + 1 < len(images) else None
        output_image_path = os.path.join(merge_output_path, f"combined_{(i // 2) + 1}.png")
        merge_images(image1, image2, output_image_path)

    print("PDF文件已转换为高质量图像并两两拼接到A4纸上。")

if __name__ == "__main__":
    main()
