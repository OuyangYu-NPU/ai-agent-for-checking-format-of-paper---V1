"""论文格式审查智能体 - 主程序"""
import os
import sys
import json
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse, Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import asyncio
from concurrent.futures import ThreadPoolExecutor
import queue
from urllib.parse import quote

# 设置日志
def setup_logging():
    """设置日志系统"""
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # 日志格式
    log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"
    
    # 主日志
    main_handler = logging.FileHandler(
        log_dir / f"paper_review_{datetime.now().strftime('%Y%m%d')}.log",
        encoding='utf-8'
    )
    main_handler.setFormatter(logging.Formatter(log_format, date_format))
    
    # PDF解析日志
    pdf_handler = logging.FileHandler(
        log_dir / f"pdf_parsing_{datetime.now().strftime('%Y%m%d')}.log",
        encoding='utf-8'
    )
    pdf_handler.setFormatter(logging.Formatter(log_format, date_format))
    
    # 审查过程日志
    review_handler = logging.FileHandler(
        log_dir / f"review_process_{datetime.now().strftime('%Y%m%d')}.log",
        encoding='utf-8'
    )
    review_handler.setFormatter(logging.Formatter(log_format, date_format))
    
    # LLM调用日志
    llm_handler = logging.FileHandler(
        log_dir / f"llm_calls_{datetime.now().strftime('%Y%m%d')}.log",
        encoding='utf-8'
    )
    llm_handler.setFormatter(logging.Formatter(log_format, date_format))
    
    # 控制台输出
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(logging.Formatter(log_format, date_format))
    
    # 配置根日志
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(main_handler)
    root_logger.addHandler(console_handler)
    
    # 配置特定日志器
    logging.getLogger("pdf_parser").addHandler(pdf_handler)
    logging.getLogger("review_process").addHandler(review_handler)
    logging.getLogger("llm_calls").addHandler(llm_handler)
    
    logger = logging.getLogger(__name__)
    logger.info("日志系统初始化完成")
    date_str = datetime.now().strftime("%Y%m%d")
    logger.info(f"  - 主日志文件: {log_dir / f'paper_review_{date_str}.log'}")
    logger.info(f"  - PDF解析日志: {log_dir / f'pdf_parsing_{date_str}.log'}")
    logger.info(f"  - 审查过程日志: {log_dir / f'review_process_{date_str}.log'}")
    logger.info(f"  - LLM调用日志: {log_dir / f'llm_calls_{date_str}.log'}")


# 初始化日志
setup_logging()
logger = logging.getLogger(__name__)

# 导入应用模块
from app.config import UPLOAD_DIR, OPENAI_API_KEY, OPENAI_BASE_URL, OPENAI_MODEL
from app.pdf_parser import PDFParser, llm_correct_extraction
from app.reviewer import PaperReviewer
from app.checklist import get_checklist_manager
from app.llm_service import get_llm_service, update_llm_config, markdown_to_text
from app.report_generator import ReportGenerator
from app.models import PaperStructure, ReviewResult

# 创建FastAPI应用
app = FastAPI(
    title="论文格式审查智能体",
    description="基于AI的学位论文格式自动审查系统",
    version="1.0.0"
)

# CORS配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 静态文件
app.mount("/static", StaticFiles(directory="templates"), name="static")


# ========== 数据模型 ==========

class APIConfig(BaseModel):
    api_key: Optional[str] = None
    base_url: Optional[str] = None
    model: Optional[str] = None


class ChecklistUpdate(BaseModel):
    item_id: str
    enabled: bool


# ========== 页面路由 ==========

@app.get("/", response_class=HTMLResponse)
async def index():
    """主页"""
    with open("templates/index.html", "r", encoding="utf-8") as f:
        return f.read()


# ========== API路由 ==========

@app.get("/api/health")
async def health_check():
    """健康检查"""
    return {"status": "ok", "timestamp": datetime.now().isoformat()}


@app.get("/api/papers")
async def list_papers():
    """获取可用的论文列表"""
    papers = []
    upload_path = Path(UPLOAD_DIR)
    
    if upload_path.exists():
        for folder in sorted(upload_path.iterdir(), reverse=True):
            if folder.is_dir() and folder.name.startswith("pdf_"):
                for pdf_file in folder.glob("*.pdf"):
                    papers.append({
                        "name": pdf_file.name,
                        "path": str(pdf_file),
                        "folder": folder.name,
                        "time": folder.name.replace("pdf_", "")
                    })
    
    return {"papers": papers}


@app.post("/api/upload")
async def upload_files(files: list[UploadFile] = File(...)):
    """上传PDF文件"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    upload_folder = Path(UPLOAD_DIR) / f"pdf_{timestamp}"
    upload_folder.mkdir(parents=True, exist_ok=True)
    
    uploaded = []
    for file in files:
        if file.filename.endswith(".pdf"):
            file_path = upload_folder / file.filename
            with open(file_path, "wb") as f:
                content = await file.read()
                f.write(content)
            uploaded.append(file.filename)
            logger.info(f"上传文件: {file.filename}")
    
    return {
        "success": True,
        "message": f"成功上传{len(uploaded)}个文件",
        "files": uploaded
    }


@app.post("/api/upload-with-progress")
async def upload_files_with_progress(files: list[UploadFile] = File(...)):
    """上传PDF文件并返回解析进度（SSE）"""
    
    # 在进入生成器之前先读取所有文件内容（避免文件被关闭的问题）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    upload_folder = Path(UPLOAD_DIR) / f"pdf_{timestamp}"
    upload_folder.mkdir(parents=True, exist_ok=True)
    
    # 先读取并保存所有文件
    uploaded = []
    file_paths = []
    for file in files:
        if file.filename.endswith(".pdf"):
            file_path = upload_folder / file.filename
            content = await file.read()
            with open(file_path, "wb") as f:
                f.write(content)
            uploaded.append(file.filename)
            file_paths.append(file_path)
            logger.info(f"上传文件: {file.filename}")
    
    async def progress_generator():
        try:
            # 步骤1: 上传完成
            yield f"data: {json.dumps({'progress': 20, 'status': f'已上传 {len(uploaded)} 个文件', 'step': 'upload'})}\n\n"
            await asyncio.sleep(0.1)
            
            # 步骤2: 解析PDF
            for idx, file_path in enumerate(file_paths):
                yield f"data: {json.dumps({'progress': 30, 'status': f'正在解析 {file_path.name}...', 'step': 'parse'})}\n\n"
                
                parser = PDFParser(str(file_path))
                if not parser.load():
                    yield f"data: {json.dumps({'error': f'无法加载PDF: {file_path.name}'})}\n\n"
                    return
                
                # 使用异步队列接收进度更新
                progress_queue = asyncio.Queue()
                loop = asyncio.get_running_loop()
                
                def parse_with_progress():
                    """在后台线程中执行解析，通过队列传递进度"""
                    def progress_callback(current, total, status):
                        # 计算进度百分比（30-60% 用于解析）
                        percent = min(30 + int((current / total) * 30), 60)
                        # 将进度放入队列（使用线程安全的方式）
                        try:
                            asyncio.run_coroutine_threadsafe(
                                progress_queue.put({
                                    'current': current,
                                    'total': total,
                                    'status': status,
                                    'percent': percent
                                }),
                                loop
                            )
                        except:
                            pass  # 如果事件循环已关闭，忽略错误
                    
                    try:
                        if not parser.parse(progress_callback=progress_callback):
                            asyncio.run_coroutine_threadsafe(
                                progress_queue.put({'error': f'PDF解析失败: {file_path.name}'}),
                                loop
                            )
                            return None
                        return parser
                    except Exception as e:
                        asyncio.run_coroutine_threadsafe(
                            progress_queue.put({'error': str(e)}),
                            loop
                        )
                        return None
                
                # 在线程池中执行解析
                with ThreadPoolExecutor() as executor:
                    parse_future = loop.run_in_executor(executor, parse_with_progress)
                    
                    # 实时接收进度更新
                    parse_done = False
                    while not parse_done:
                        # 检查是否有进度更新（非阻塞）
                        try:
                            progress_data = await asyncio.wait_for(progress_queue.get(), timeout=0.1)
                            if 'error' in progress_data:
                                yield f"data: {json.dumps({'error': progress_data['error']})}\n\n"
                                return
                            yield f"data: {json.dumps({'progress': progress_data['percent'], 'status': progress_data['status'], 'step': 'parse'})}\n\n"
                        except asyncio.TimeoutError:
                            pass
                        
                        # 检查解析是否完成
                        if parse_future.done():
                            parse_done = True
                    
                    # 获取解析结果
                    parser = await parse_future
                    if parser is None:
                        return
                
                yield f"data: {json.dumps({'progress': 60, 'status': '正在提取结构信息...', 'step': 'extract'})}\n\n"
                await asyncio.sleep(0.1)
                
                # 步骤3: 提取论文信息
                structure = parser.extract_paper_info()
                
                yield f"data: {json.dumps({'progress': 75, 'status': '正在转换为JSON格式...', 'step': 'extract'})}\n\n"
                await asyncio.sleep(0.1)
                
                # 转换为JSON
                json_output = file_path.parent / f"{file_path.stem}.json"
                try:
                    parser.convert_to_json(str(json_output))
                except Exception as e:
                    logger.warning(f"JSON转换失败: {e}")
                
                yield f"data: {json.dumps({'progress': 85, 'status': '正在保存解析结果...', 'step': 'save'})}\n\n"
                await asyncio.sleep(0.1)
                
                # 步骤4: 保存结构文件
                structure_file = file_path.parent / f"{file_path.stem}_structure.json"
                with open(structure_file, 'w', encoding='utf-8') as f:
                    json.dump(structure.model_dump(), f, ensure_ascii=False, indent=2)
                
                parser.close()
                
                yield f"data: {json.dumps({'progress': 95, 'status': '正在完成...', 'step': 'save'})}\n\n"
                await asyncio.sleep(0.1)
            
            # 完成
            yield f"data: {json.dumps({'progress': 100, 'status': '解析完成！', 'step': 'save', 'success': True, 'message': f'成功上传并解析 {len(uploaded)} 个文件', 'files': uploaded})}\n\n"
            
        except Exception as e:
            logger.error(f"上传处理错误: {e}")
            import traceback
            logger.error(traceback.format_exc())
            yield f"data: {json.dumps({'error': str(e)})}\n\n"
    
    return StreamingResponse(
        progress_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )


@app.get("/api/paper-info")
async def get_paper_info(paper_path: str):
    """获取论文信息"""
    pdf_path = Path(paper_path)
    if not pdf_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    # 检查是否有已保存的结构文件
    structure_file = pdf_path.parent / f"{pdf_path.stem}_structure.json"
    
    if structure_file.exists():
        # 加载已保存的结构
        with open(structure_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        structure = PaperStructure(**data)
        logger.info(f"加载已保存的论文结构: {structure_file}")
    else:
        # 解析PDF
        parser = PDFParser(paper_path)
        if not parser.load() or not parser.parse():
            raise HTTPException(status_code=500, detail="PDF解析失败")
        
        structure = parser.extract_paper_info()
        parser.close()
        
        # 保存结构文件
        with open(structure_file, 'w', encoding='utf-8') as f:
            json.dump(structure.model_dump(), f, ensure_ascii=False, indent=2)
        logger.info(f"保存论文结构: {structure_file}")
    
    # 构建返回信息
    info = {
        "title": structure.title,
        "title_en": structure.title_en,
        "author": structure.author,
        "student_id": structure.student_id,
        "college": structure.college,
        "discipline": structure.discipline,
        "supervisor": structure.supervisor,
        "defense_date": structure.defense_date,
        "degree_type": structure.degree_type,
        "total_pages": structure.total_pages,
        "chapters": [ch.model_dump() for ch in structure.chapters],
        "abstract_cn": structure.abstract_cn[:500] if structure.abstract_cn else None,
        "abstract_en": structure.abstract_en[:500] if structure.abstract_en else None,
        "keywords_cn": structure.keywords_cn,
        "keywords_en": structure.keywords_en,
        "references_count": len(structure.references),
        "sections": [s.model_dump() for s in structure.sections]
    }
    
    statistics = {
        "word_count": structure.word_count,
        "char_count": structure.char_count,
        "image_count": len(structure.figures),
        "table_count": len(structure.tables),
        "reference_count": len(structure.references)
    }
    
    return {
        "success": True,
        "info": info,
        "statistics": statistics
    }


@app.get("/api/chapter-content")
async def get_chapter_content(paper_path: str, start_page: int, end_page: int):
    """获取章节全文内容（从JSON文件的markdown数据）"""
    try:
        # 修复路径：处理URL编码和路径分隔符
        import urllib.parse
        
        # 先解码URL编码
        decoded_path = urllib.parse.unquote(paper_path)
        logger.info(f"原始路径参数: {paper_path}")
        logger.info(f"解码后路径: {decoded_path}")
        
        # 尝试不同的路径格式
        pdf_path = None
        path_variants = [
            Path(decoded_path),  # 直接使用解码后的路径
            Path(decoded_path.replace('\\', '/')),  # 统一为正斜杠
            Path(decoded_path.replace('/', '\\')),  # 统一为反斜杠（Windows）
            Path(paper_path),  # 原始路径（如果未编码）
        ]
        
        # 如果路径看起来不完整（缺少目录分隔符），尝试在uploads目录下查找
        if not any(p.exists() for p in path_variants):
            # 提取文件名（可能包含部分路径信息）
            path_parts = decoded_path.replace('\\', '/').split('/')
            filename = path_parts[-1] if path_parts else decoded_path
            
            # 如果文件名看起来不完整（可能路径被错误拼接），尝试提取实际文件名
            # 例如：uploadspdf_20251203_131055【博士学位论文】-建筑-1.pdf
            # 应该找到：uploads/pdf_20251203_131055/【博士学位论文】-建筑-1.pdf
            if 'pdf_' in filename and not filename.startswith('uploads'):
                # 尝试从文件名中提取文件夹名和实际文件名
                # 格式可能是：pdf_20251203_131055【文件名】.pdf
                import re
                match = re.match(r'pdf_(\d{8}_\d{6})(.+)', filename)
                if match:
                    folder_name = f"pdf_{match.group(1)}"
                    actual_filename = match.group(2)
                    potential_path = Path("uploads") / folder_name / actual_filename
                    if potential_path.exists():
                        pdf_path = potential_path
                        logger.info(f"通过模式匹配找到文件: {pdf_path}")
            
            # 如果还没找到，尝试在uploads目录下查找
            if not pdf_path:
                uploads_dir = Path("uploads")
                if uploads_dir.exists():
                    # 查找所有子目录中的PDF文件
                    for subdir in uploads_dir.iterdir():
                        if subdir.is_dir():
                            pdf_file = subdir / filename
                            if pdf_file.exists():
                                pdf_path = pdf_file
                                logger.info(f"在uploads目录下找到文件: {pdf_path}")
                                break
                    
                    # 如果还是没找到，尝试模糊匹配文件名
                    if not pdf_path:
                        for subdir in uploads_dir.iterdir():
                            if subdir.is_dir():
                                for pdf_file in subdir.glob("*.pdf"):
                                    # 检查文件名是否相似（忽略路径前缀）
                                    if pdf_file.name == filename or filename in pdf_file.name or pdf_file.name in filename:
                                        pdf_path = pdf_file
                                        logger.info(f"通过模糊匹配找到文件: {pdf_path}")
                                        break
                                if pdf_path:
                                    break
        
        # 再次尝试所有路径变体
        if not pdf_path:
            for path_variant in path_variants:
                if path_variant.exists():
                    pdf_path = path_variant
                    logger.info(f"找到文件: {pdf_path}")
                    break
        
        # 如果找不到PDF文件，直接尝试查找JSON文件
        json_file = None
        if pdf_path and pdf_path.exists():
            json_file = pdf_path.parent / f"{pdf_path.stem}.json"
        else:
            # 直接查找JSON文件
            logger.info("PDF文件未找到，尝试直接查找JSON文件")
            # 提取可能的文件名
            filename_base = Path(decoded_path).stem
            # 在uploads目录下查找JSON文件
            uploads_dir = Path("uploads")
            if uploads_dir.exists():
                # 查找所有JSON文件
                for json_file_candidate in uploads_dir.rglob("*.json"):
                    # 检查文件名是否匹配（忽略路径）
                    if filename_base in json_file_candidate.stem or json_file_candidate.stem in filename_base:
                        json_file = json_file_candidate
                        logger.info(f"找到匹配的JSON文件: {json_file}")
                        break
                
                # 如果还是没找到，查找最近修改的JSON文件
                if not json_file:
                    json_files = sorted(uploads_dir.rglob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
                    if json_files:
                        json_file = json_files[0]
                        logger.info(f"使用最近修改的JSON文件: {json_file}")
        
        if not json_file or not json_file.exists():
            logger.error(f"JSON文件不存在")
            if pdf_path:
                logger.error(f"PDF路径: {pdf_path}")
                logger.error(f"尝试的JSON路径: {pdf_path.parent / f'{pdf_path.stem}.json'}")
            # 列出所有可用的JSON文件
            uploads_dir = Path("uploads")
            if uploads_dir.exists():
                all_json_files = list(uploads_dir.rglob("*.json"))
                logger.info(f"uploads目录下的所有JSON文件: {[str(f) for f in all_json_files[:10]]}")
            raise HTTPException(status_code=404, detail=f"JSON文件不存在，请确保PDF已解析")
        
        try:
            logger.info(f"读取JSON文件: {json_file}")
            with open(json_file, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            pages = json_data.get("pages", [])
            total_pages = json_data.get("total_pages", len(pages))
            logger.info(f"JSON文件包含 {len(pages)} 页数据，总页数: {total_pages}")
            logger.info(f"请求页面范围: {start_page + 1} - {end_page + 1} (0-based: {start_page} - {end_page})")
            
            page_contents = []
            
            # 提取指定页面的markdown内容
            # start_page和end_page是0-based索引，需要转换为1-based的page_number
            for page_idx in range(start_page, min(end_page + 1, total_pages)):
                # page_number是1-based，所以page_idx + 1才是page_number
                target_page_number = page_idx + 1
                
                # 查找对应的页面数据
                page_data = None
                for p in pages:
                    if p.get("page_number") == target_page_number:
                        page_data = p
                        break
                
                if page_data:
                    # 优先使用body部分的markdown，但需要过滤掉页眉页脚
                    body_data = page_data.get("body", {})
                    markdown_content = body_data.get("markdown", "")
                    
                    # 如果没有markdown，使用raw_text
                    if not markdown_content:
                        markdown_content = body_data.get("raw_text", "")
                    
                    # 如果还是没有，尝试从elements构建（只使用position为body的元素）
                    if not markdown_content:
                        elements = body_data.get("elements", [])
                        text_parts = []
                        for elem in elements:
                            # 只提取position为"body"的文本元素，过滤掉header和footer
                            elem_position = elem.get("position", "").lower()
                            if elem.get("type") == "text" and elem_position != "header" and elem_position != "footer":
                                text_parts.append(elem.get("raw_text", ""))
                        markdown_content = " ".join(text_parts)
                    
                    # 清理可能残留的页眉页脚内容
                    if markdown_content:
                        # 移除常见的页眉页脚模式
                        import re
                        # 移除页眉常见格式（如"XXXXX 硕士学位论文"、"西安建筑科技大学硕士学位论文"等）
                        markdown_content = re.sub(r'.*?硕士学位论文.*?\n', '', markdown_content, flags=re.IGNORECASE)
                        markdown_content = re.sub(r'.*?博士学位论文.*?\n', '', markdown_content, flags=re.IGNORECASE)
                        markdown_content = re.sub(r'.*?西安建筑科技大学.*?\n', '', markdown_content, flags=re.IGNORECASE)
                        
                        # 移除页脚常见格式（页码、短文本行等）
                        lines = markdown_content.split('\n')
                        cleaned_lines = []
                        for line in lines:
                            line = line.strip()
                            if not line:
                                continue
                            
                            # 跳过看起来像页码的行（只匹配明显的页码格式，避免误删正文）
                            # 罗马数字页码（单独一行，只有罗马数字和分隔符）
                            if re.match(r'^[\s\-—\[\]()]*[IVXLCDMivxlcdm]+[\s\-—\[\]()]*$', line):
                                continue
                            # 纯数字页码（单独一行，只有数字和分隔符，且数字较小，通常是页码）
                            if re.match(r'^[\s\-—\[\]()]*\d{1,3}[\s\-—\[\]()]*$', line) and len(line) < 20:
                                continue
                            # 分数格式页码（如"1/10"）
                            if re.match(r'^\d+\s*/\s*\d+$', line):
                                continue
                            # 中文页码格式（如"第1页"）
                            if re.match(r'^第\s*[\d一二三四五六七八九十百]+\s*页\s*$', line):
                                continue
                            # Page格式（如"Page 1"）
                            if re.match(r'^page\s*\d+\s*$', line, re.IGNORECASE):
                                continue
                            
                            # 跳过太短的行（可能是页眉页脚，但保留有意义的短行）
                            if len(line) < 2:
                                continue
                            
                            cleaned_lines.append(line)
                        
                        markdown_content = '\n'.join(cleaned_lines).strip()
                    
                    if markdown_content:
                        page_contents.append({
                            "page_number": target_page_number,
                            "markdown": markdown_content
                        })
                        logger.info(f"找到第{target_page_number}页的内容，长度: {len(markdown_content)} 字符")
                    else:
                        logger.warning(f"第{target_page_number}页的body中没有内容")
                else:
                    logger.warning(f"未找到第{target_page_number}页的数据")
            
            if not page_contents:
                logger.error(f"未找到任何页面内容，请求范围: {start_page + 1} - {end_page + 1}")
                raise HTTPException(status_code=404, detail=f"未找到指定页面的内容（第{start_page + 1}-{end_page + 1}页）")
            
            logger.info(f"成功提取 {len(page_contents)} 页的内容")
            
            # 合并所有页面的markdown内容
            full_markdown = "\n\n---\n\n".join([
                f"## 第{pc['page_number']}页\n\n{pc['markdown']}" 
                for pc in page_contents
            ])
            
            return {
                "success": True,
                "content": full_markdown,
                "format": "markdown",
                "page_count": len(page_contents),
                "start_page": start_page + 1,
                "end_page": end_page + 1
            }
            
        except Exception as e:
            logger.error(f"从JSON读取页面内容失败: {e}")
            import traceback
            logger.error(traceback.format_exc())
            raise HTTPException(status_code=500, detail=f"读取JSON文件失败: {str(e)}")
            
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"获取章节内容失败: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"获取章节内容失败: {str(e)}")


@app.post("/api/review")
async def review_paper(
    paper_path: str = Form(...),
    use_ai: bool = Form(True),
    degree_type: str = Form("phd")
):
    """审查论文"""
    pdf_path = Path(paper_path)
    if not pdf_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    logger.info("=" * 60)
    logger.info(f"开始审查论文: {pdf_path.name}")
    logger.info(f"  学位类型: {degree_type}")
    logger.info(f"  使用AI: {use_ai}")
    
    # 检查是否有已保存的结构文件
    structure_file = pdf_path.parent / f"{pdf_path.stem}_structure.json"
    structure = None
    
    if structure_file.exists():
        with open(structure_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        structure = PaperStructure(**data)
        logger.info(f"加载已保存的论文结构")
    
    # 解析PDF
    parser = PDFParser(paper_path)
    if not parser.load() or not parser.parse():
        raise HTTPException(status_code=500, detail="PDF解析失败")
    
    # 如果没有已保存的结构，提取并可能进行LLM矫正
    if not structure:
        structure = parser.extract_paper_info()
        
        # LLM矫正（条件触发）
        if use_ai:
            llm_service = get_llm_service()
            if llm_service.is_available():
                structure = await llm_correct_extraction(structure, llm_service)
        
        # 保存结构文件
        with open(structure_file, 'w', encoding='utf-8') as f:
            json.dump(structure.model_dump(), f, ensure_ascii=False, indent=2)
    
    # 创建审查器（传入已处理的structure）
    reviewer = PaperReviewer(parser, degree_type=degree_type, structure=structure)
    result = reviewer.review()
    
    # AI深度分析
    if use_ai:
        llm_service = get_llm_service()
        if llm_service.is_available():
            try:
                logger.info("执行AI综合审查...")
                ai_result = await llm_service.comprehensive_review(structure, degree_type)
                
                # 处理AI结果
                ai_analysis_parts = []
                
                if ai_result and ai_result.get("typos"):
                    ai_analysis_parts.append("【错别字问题】")
                    for item in ai_result["typos"][:5]:
                        location = item.get('location', '')
                        original = item.get('original', '')
                        suggestion = item.get('suggestion', '')
                        # 确保location中包含页码信息，如果没有则尝试提取
                        if location and '第' not in location and '页' not in location:
                            # 如果没有页码信息，保持原样，让格式化方法处理
                            ai_analysis_parts.append(f"  - {location}: {original} → {suggestion}")
                        else:
                            ai_analysis_parts.append(f"  - {location}: {original} → {suggestion}")
                
                if ai_result and ai_result.get("writing_style"):
                    ai_analysis_parts.append("\n【学术写作规范】")
                    for item in ai_result["writing_style"][:5]:
                        location = item.get('location', '')
                        issue = item.get('issue', '')
                        suggestion = item.get('suggestion', '')
                        if location:
                            ai_analysis_parts.append(f"  - {location}: {issue}")
                        else:
                            ai_analysis_parts.append(f"  - {issue}")
                        if suggestion:
                            ai_analysis_parts.append(f"    建议: {suggestion}")
                
                if ai_result and ai_result.get("content_quality"):
                    ai_analysis_parts.append("\n【内容质量】")
                    for item in ai_result["content_quality"][:5]:
                        location = item.get('location', '')
                        issue = item.get('issue', '')
                        suggestion = item.get('suggestion', '')
                        if location:
                            ai_analysis_parts.append(f"  - {location}: {issue}")
                        else:
                            ai_analysis_parts.append(f"  - {issue}")
                        if suggestion:
                            ai_analysis_parts.append(f"    建议: {suggestion}")
                
                if ai_result and ai_result.get("overall_comments"):
                    ai_analysis_parts.append(f"\n【总体评价】\n{ai_result['overall_comments']}")
                
                result.ai_analysis = "\n".join(ai_analysis_parts) if ai_analysis_parts else None
            except Exception as e:
                logger.error(f"AI审查失败: {e}")
                import traceback
                logger.error(traceback.format_exc())
                result.ai_analysis = f"AI审查过程中出现错误: {str(e)}"
        else:
            logger.warning("LLM服务不可用，跳过AI审查")
            result.ai_analysis = "LLM服务未配置，无法进行AI深度分析"
    
    parser.close()
    
    # 转换summary为纯文本
    if result.summary:
        result.summary = markdown_to_text(result.summary)
    
    # 保存审查结果
    review_result_file = pdf_path.parent / f"{pdf_path.stem}_review.json"
    with open(review_result_file, 'w', encoding='utf-8') as f:
        json.dump(result.model_dump(), f, ensure_ascii=False, indent=2)
    logger.info(f"审查结果已保存: {review_result_file}")
    
    logger.info(f"审查完成: 得分={result.overall_score}, 问题数={result.total_issues}")
    logger.info("=" * 60)
    
    return {
        "success": True,
        "result": result.model_dump()
    }


@app.get("/api/export-pdf-report")
async def export_pdf_report(paper_path: str):
    """导出PDF格式的审查报告"""
    pdf_path = Path(paper_path)
    if not pdf_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    # 加载审查结果
    review_result_file = pdf_path.parent / f"{pdf_path.stem}_review.json"
    if not review_result_file.exists():
        raise HTTPException(status_code=404, detail="审查结果不存在，请先进行审查")
    
    with open(review_result_file, 'r', encoding='utf-8') as f:
        review_data = json.load(f)
    result = ReviewResult(**review_data)
    
    # 加载论文结构
    structure_file = pdf_path.parent / f"{pdf_path.stem}_structure.json"
    if not structure_file.exists():
        raise HTTPException(status_code=404, detail="论文结构文件不存在")
    
    with open(structure_file, 'r', encoding='utf-8') as f:
        structure_data = json.load(f)
    structure = PaperStructure(**structure_data)
    
    # 生成PDF报告
    try:
        generator = ReportGenerator(structure, result)
        pdf_data = generator.generate_pdf_report()
        
        # 处理中文文件名编码问题
        # 提取文件名（去除路径和扩展名）
        stem = pdf_path.stem
        
        # 创建ASCII安全的文件名（移除或替换非ASCII字符）
        stem_ascii = stem.encode('ascii', 'ignore').decode('ascii')
        if not stem_ascii:
            stem_ascii = "paper"  # 如果全是中文，使用默认名称
        
        # 只使用ASCII文件名，避免编码问题
        filename = f"{stem_ascii}_review_report.pdf"
        
        return Response(
            content=pdf_data,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )
    except ImportError as e:
        raise HTTPException(status_code=500, detail=f"PDF生成功能不可用: {str(e)}")
    except Exception as e:
        logger.error(f"生成PDF报告失败: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"生成PDF报告失败: {str(e)}")


@app.post("/api/generate-report")
async def generate_report(paper_path: str = Form(...)):
    """生成改进报告"""
    pdf_path = Path(paper_path)
    if not pdf_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    # 加载结构
    structure_file = pdf_path.parent / f"{pdf_path.stem}_structure.json"
    if structure_file.exists():
        with open(structure_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        structure = PaperStructure(**data)
    else:
        parser = PDFParser(paper_path)
        if not parser.load() or not parser.parse():
            raise HTTPException(status_code=500, detail="PDF解析失败")
        structure = parser.extract_paper_info()
        parser.close()
    
    # 生成报告
    llm_service = get_llm_service()
    if llm_service.is_available():
        report = await llm_service.generate_improvement_report(structure, [])
        return {"success": True, "report": report}
    else:
        return {"success": False, "message": "LLM服务不可用"}


# ========== 配置API ==========

@app.get("/api/config")
async def get_config():
    """获取API配置"""
    return {
        "success": True,
        "config": {
            "has_key": bool(OPENAI_API_KEY),
            "base_url": OPENAI_BASE_URL,
            "model": OPENAI_MODEL
        }
    }


@app.post("/api/config")
async def save_config(config: APIConfig):
    """保存API配置"""
    # 更新环境变量
    if config.api_key:
        os.environ["OPENAI_API_KEY"] = config.api_key
    if config.base_url:
        os.environ["OPENAI_BASE_URL"] = config.base_url
    if config.model:
        os.environ["OPENAI_MODEL"] = config.model
    
    # 更新LLM服务
    update_llm_config(
        api_key=config.api_key or OPENAI_API_KEY,
        base_url=config.base_url or OPENAI_BASE_URL,
        model=config.model or OPENAI_MODEL
    )
    
    return {"success": True, "message": "配置已保存"}


@app.post("/api/config/test")
async def test_config(config: APIConfig):
    """测试API配置"""
    from app.llm_service import LLMService
    
    service = LLMService(
        api_key=config.api_key or OPENAI_API_KEY,
        base_url=config.base_url or OPENAI_BASE_URL,
        model=config.model or OPENAI_MODEL
    )
    
    if await service.test_connection():
        return {"success": True, "message": "连接成功"}
    else:
        return {"success": False, "message": "连接失败，请检查配置"}


# ========== 检查清单API ==========

@app.get("/api/checklist")
async def get_checklist():
    """获取检查清单"""
    manager = get_checklist_manager()
    config = manager.get_config()
    
    return {
        "success": True,
        "checklist": {
            "name": config.name,
            "version": config.version,
            "categories": config.categories,
            "items": [item.model_dump() for item in config.items]
        }
    }


@app.post("/api/checklist/update")
async def update_checklist_item(update: ChecklistUpdate):
    """更新检查项"""
    manager = get_checklist_manager()
    success = manager.update_item(update.item_id, enabled=update.enabled)
    
    return {"success": success}


@app.post("/api/checklist/reset")
async def reset_checklist():
    """重置检查清单"""
    manager = get_checklist_manager()
    manager.reset_to_default()
    
    return {"success": True, "message": "已重置为默认配置"}


@app.get("/api/checklist/export")
async def export_checklist():
    """导出检查清单"""
    manager = get_checklist_manager()
    data = manager.export_config()
    
    return {"success": True, "data": data}


@app.post("/api/checklist/import")
async def import_checklist(file: UploadFile = File(...)):
    """导入检查清单"""
    content = await file.read()
    data = json.loads(content.decode('utf-8'))
    
    manager = get_checklist_manager()
    success = manager.import_config(data)
    
    return {"success": success, "message": "导入成功" if success else "导入失败"}


# ========== 启动 ==========

if __name__ == "__main__":
    import uvicorn
    
    logger.info("启动论文格式审查智能体...")
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )

