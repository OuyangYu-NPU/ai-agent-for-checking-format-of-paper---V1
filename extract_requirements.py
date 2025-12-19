# -*- coding: utf-8 -*-
"""提取论文格式要求"""
import fitz
import os
import json
from pathlib import Path

def extract_pdf_text():
    """从PDF中提取文本"""
    requirements_dir = Path(__file__).parent / "要求"
    
    if not requirements_dir.exists():
        print(f"目录不存在: {requirements_dir}")
        return
    
    for pdf_file in requirements_dir.glob("*.pdf"):
        print(f"正在处理: {pdf_file.name}")
        try:
            doc = fitz.open(str(pdf_file))
            text = ""
            for page in doc:
                text += page.get_text()
            
            # 保存文本
            output_file = requirements_dir / f"{pdf_file.stem}.txt"
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(text)
            
            print(f"已保存到: {output_file}")
            print(f"\n前5000字符:\n{text[:5000]}")
            
        except Exception as e:
            print(f"处理失败: {e}")

if __name__ == "__main__":
    extract_pdf_text()

