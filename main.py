import sys
import io
import contextlib
import re  
import openai
import ssl  
import os

# ─────────────────────────────────────────────
# 1. 依赖检查层
# ─────────────────────────────────────────────
try:
    import pptx  
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError as e:
    print(f"❌ 运行环境库缺失: {e}，请执行 pip install python-pptx")
    sys.exit(1)

# ─────────────────────────────────────────────
# 2. 工具层：文件解析与执行沙盒（必须在类定义前）
# ─────────────────────────────────────────────
def extract_text_from_file(file_path: str) -> str:
    """提取本地文件的纯文本内容，支持 txt 和 docx"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.txt':
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        except Exception as e:
            return f"❌ 读取文本文件失败: {e}"
    elif ext == '.docx':
        try:
            import docx
            doc = docx.Document(file_path)
            return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        except ImportError:
            return "❌ 错误：未安装 python-docx 库，请执行 pip install python-docx。"
        except Exception as e:
            return f"❌ 解析 Word 文件失败: {e}"
    else:
        return "❌ 错误：不支持的文件格式。仅支持 .txt 和 .docx。"

def execute_python_code(code_string: str) -> str:
    """接收代码并执行。在执行前强行剔除可能导致崩溃的引用"""
    # 补丁：强行删除所有包含 MSO_ANCHOR 的行，这是 python-pptx 在某些环境下容易报错的点
    code_string = re.sub(r'.*MSO_ANCHOR.*', '', code_string)
    
    output_buffer = io.StringIO()
    try:
        ssl._create_default_https_context = ssl._create_unverified_context
        with contextlib.redirect_stdout(output_buffer):
            # 预注入常用模块与对象，确保模型生成的代码能直接访问
            exec_globals = {
                'pptx': pptx,
                'Inches': Inches, 
                'Pt': Pt, 
                'RGBColor': RGBColor, 
                'PP_ALIGN': PP_ALIGN,
                'Presentation': pptx.Presentation
            }
            exec(code_string, exec_globals)
        return output_buffer.getvalue()
    except BaseException as e: 
        return f"代码执行报错: {type(e).__name__}: {str(e)}"

# ─────────────────────────────────────────────
# 3. 规划与执行层：学术风 PPT 引擎
# ─────────────────────────────────────────────
class AutoPPTAgent:
    def __init__(self):
        # 请确保 API Key 有效
        self.client = openai.OpenAI(
            api_key="请输入deepseek的API Key", 
            base_url="https://api.deepseek.com"
        )
        
    def _generate_execution_plan(self, user_input: str) -> str:
        """阶段 1：规划大纲"""
        meta_prompt = f"""
        你是一个严谨的学术导师。针对以下需求规划一份结构严谨的学术 PPT 大纲。
        【结构要求】：
        1. 必须包含封面页（标题/副标题）、目录、核心研究内容页（3-5页）、总结页。
        2. 每页需标明标题、一级要点（加粗）、二级补充内容。
        3. 逻辑必须连贯，学术术语使用准确。
        
        【用户输入】：
        "{user_input}"
        """
        response = self.client.chat.completions.create(
            model="deepseek-chat",
            messages=[{"role": "system", "content": meta_prompt}],
        )
        return response.choices[0].message.content

    def _generate_and_run_code(self, execution_plan: str, max_retries: int = 2) -> tuple:
        """阶段 2：编写代码"""
        coder_system_prompt = """
        你是一个 Python 自动化工程师，专精于 python-pptx。
        【强制约束】：
        1. 布局：必须使用 `prs.slide_layouts[6]` (空白布局)。
        2. 样式：内容页顶部必须有一个深蓝矩形 (0, 51, 102) 作为标题背景，高度 0.8 英寸。
        3. 字体：全篇使用 'Arial' 或 'Times New Roman'。标题白色 26pt。
        4. 安全：禁止使用 MSO_ANCHOR。所有文本框必须通过 prs.slides[i].shapes.add_textbox() 显式指定 Inches 位置。
        5. 结尾：必须执行 `prs.save("output_presentation.pptx")` 并在代码最后 print("✅ PPT生成成功")。
        """
        
        messages = [
            {"role": "system", "content": coder_system_prompt},
            {"role": "user", "content": f"根据以下大纲生成完整 Python 代码：\n{execution_plan}"}
        ]
        
        python_code = ""
        execution_result = ""

        for attempt in range(max_retries + 1):
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=messages,
            )
            full_response = response.choices[0].message.content
            
            # 提取 Markdown 里的 Python 代码块
            code_match = re.search(r"```python\n(.*?)\n```", full_response, re.DOTALL)
            python_code = code_match.group(1) if code_match else full_response
                
            execution_result = execute_python_code(python_code)
            
            if "✅ PPT生成成功" in execution_result:
                return python_code, execution_result 
                
            print(f"\n[⚠️ 修复尝试] 第 {attempt+1} 次重试...")
            messages.append({"role": "assistant", "content": full_response})
            messages.append({"role": "user", "content": f"代码执行失败信息：{execution_result}。请重新检查坐标计算和属性引用，确保不使用任何枚举常量。"})
                
        return python_code, execution_result

    def generate_ppt(self, requirements: str, file_path: str = None):
        final_prompt = requirements

        # 逻辑：文件解析
        if file_path:
            # 清理路径字符
            clean_path = file_path.strip().strip("'").strip('"').strip()
            if os.path.exists(clean_path):
                print(f"0. [输入流解析] 正在读取文件: {os.path.basename(clean_path)}...\n")
                raw_text = extract_text_from_file(clean_path)
                if "❌" in raw_text:
                    print(raw_text)
                    return
                final_prompt = f"【用户要求】\n{requirements}\n\n【参考文档文本】\n{raw_text[:6000]}"
            else:
                print(f"⚠️ 找不到文件: {clean_path}，将仅按文本要求生成。")

        print("1. [文案引擎] 构建结构化学术大纲...\n")
        execution_plan = self._generate_execution_plan(final_prompt)
        
        print("2. [渲染引擎] 正在编写并运行 PPT 生成代码...\n")
        python_code, execution_result = self._generate_and_run_code(execution_plan)
        
        if "✅ PPT生成成功" in execution_result:
            print("--- 最终生成的业务逻辑代码 ---")
            print(python_code)
            print("------------------------------\n")
            print("4. [汇报总结] PPT 已成功生成并保存为：output_presentation.pptx")
        else:
            print(f"\n❌ 最终失败。错误详情:\n{execution_result}")

# ─────────────────────────────────────────────
# 4. 程序入口
# ─────────────────────────────────────────────
if __name__ == "__main__":
    agent = AutoPPTAgent()
    print("===========================================")
    print("🎓 学术 PPT Agent")
    print("===========================================")
    
    choice = input("\n是否参考本地文件？(y/n) > ").strip().lower()
    
    if choice in ['y', 'yes', '是', '1']:
        path_input = input("请拖入文件或输入路径 (.txt/.docx) > ").strip()
        user_req = input("针对该文件有什么具体定制要求？(直接按回车则默认总结生成) > ").strip()
        if not user_req: user_req = "请根据文档内容生成一份结构清晰的学术汇报PPT。"
        agent.generate_ppt(requirements=user_req, file_path=path_input)
    else:
        user_req = input("请输入 PPT 的主题或详细要求 > ").strip()
        if user_req:
            agent.generate_ppt(requirements=user_req)
        else:
            print("❌ 未输入有效需求，程序退出。")