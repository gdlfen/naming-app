import streamlit as st
import os
import re
import json
import io
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from openai import OpenAI

try:
    from lunar_python import Solar, Lunar
    HAS_LUNAR = True
except ImportError:
    HAS_LUNAR = False

# ================= 页面基础设置 =================
st.set_page_config(page_title="名字通（国学起名系统）", layout="centered", page_icon="🌟")

# 注入自定义 CSS：修改标题大小颜色，以及标签橙色底色
st.markdown("""
<style>
/* 1. 大标题缩小并改为金色 */
.main-title {
    font-size: 26px !important;
    color: #DAA520 !important; /* 稳重的金黄色 (GoldenRod) */
    font-weight: bold;
    text-align: center;
    margin-bottom: 0px;
}
.sub-title-desc {
    text-align: center;
    color: gray;
    font-size: 14px;
    margin-bottom: 20px;
}
/* 2. 小标题缩小并改为蓝色 */
.sub-title {
    font-size: 18px !important;
    color: #1E90FF !important; /* 亮蓝色 */
    font-weight: bold;
    margin-top: 15px;
    margin-bottom: 15px;
    border-left: 4px solid #1E90FF;
    padding-left: 8px;
}
/* 3. 多选框(风格偏好)改为橙底色 */
span[data-baseweb="tag"] {
    background-color: #FF8C00 !important; /* 深橙色 */
    color: white !important;
}
/* 适配橙年底色，将标签上的关闭按钮改成白色 */
span[data-baseweb="tag"] svg {
    fill: white !important;
}
</style>
""", unsafe_allow_html=True)

# ================= 会话状态初始化 =================
if 'admin_logged_in' not in st.session_state:
    st.session_state.admin_logged_in = False

if 'api_key' not in st.session_state:
    st.session_state.api_key = ""
if 'api_base' not in st.session_state:
    st.session_state.api_base = "https://api.openai.com/v1"
if 'model_name' not in st.session_state:
    st.session_state.model_name = "gpt-4o"
if 'usage_count' not in st.session_state:
    st.session_state.usage_count = 0  # 记录使用次数统计

# ================= 后台管理 (侧边栏) =================
with st.sidebar:
    st.header("⚙️ 后台管理中心")
    if not st.session_state.admin_logged_in:
        pwd = st.text_input("请输入管理员密码", type="password")
        if st.button("登录后台"):
            if pwd == "888888":  # 这里可以修改您的专属后台密码
                st.session_state.admin_logged_in = True
                st.success("登录成功！")
                st.rerun()
            else:
                st.error("密码错误！")
    else:
        st.success("管理员已登录")
        
        # --- 新增：使用情况统计 ---
        st.markdown("### 📊 使用情况 (云端实例)")
        st.metric(label="累计成功推演生成次数", value=st.session_state.usage_count)
        st.divider()
        
        # --- 优化：AI模型下拉菜单 ---
        st.markdown("### 🤖 AI 接口配置")
        st.session_state.api_base = st.text_input("接口地址 (Base URL)", value=st.session_state.api_base)
        st.session_state.api_key = st.text_input("API 密钥 (API Key)", value=st.session_state.api_key, type="password")
        
        common_models = [
            "gpt-4o", "gpt-3.5-turbo", 
            "deepseek-chat", "deepseek-reasoner", 
            "qwen-plus", "qwen-max", 
            "moonshot-v1-8k", 
            "自定义其他模型..."
        ]
        curr_model = st.session_state.model_name
        
        try:
            idx = common_models.index(curr_model)
        except ValueError:
            idx = len(common_models) - 1
            
        selected_model = st.selectbox("选择模型名称 (Model)", common_models, index=idx)
        
        if selected_model == "自定义其他模型...":
            st.session_state.model_name = st.text_input("手动输入模型名称", value=curr_model)
        else:
            st.session_state.model_name = selected_model
        
        if st.button("退出登录"):
            st.session_state.admin_logged_in = False
            st.rerun()

# ================= 前端主界面 (用户端) =================
# 替换原有的原生 title，使用注入样式的 HTML 渲染大标题
st.markdown('<div class="main-title">🌟 名字通 - 国学起名系统 v3.0</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title-desc">融合传统周易八卦、五行生克与现代 AI 智能的前沿起名引擎。</div>', unsafe_allow_html=True)

# 1. 基础信息填写
st.markdown('<div class="sub-title">1. 基础信息填写</div>', unsafe_allow_html=True)
mode = st.radio("选择起名类型", ["为人起名", "为公司起名"], horizontal=True)

col1, col2 = st.columns(2)
with col1:
    cal_type = st.selectbox("历法选择", ["公历", "农历"])
    year = st.number_input("出生年", min_value=1920, max_value=2040, value=1990)
    month = st.number_input("出生月", min_value=1, max_value=12, value=1)
with col2:
    day = st.number_input("出生日", min_value=1, max_value=31, value=1)
    hour = st.selectbox("时辰", [f"{h} 时" for h in range(24)] + ["不清楚时辰"], index=12)
    gender = st.radio("性别/属性", ["男 (乾造)", "女 (坤造)"], horizontal=True)

# 专属字段
surname, birth_place = "", ""
comp_prefix, comp_suffix, comp_industry = "", "", ""

if mode == "为人起名":
    c1, c2 = st.columns(2)
    with c1: surname = st.text_input("姓氏 (必填)", value="李")
    with c2: birth_place = st.text_input("出生地 (选填)")
else:
    c1, c2 = st.columns(2)
    with c1: comp_prefix = st.text_input("公司前缀 (如: XX市)")
    with c2: comp_suffix = st.text_input("公司后缀 (如: 科技有限公司)")
    comp_industry = st.text_input("行业/业务范围", value="互联网科技")

# 2. 风格偏好与要求
st.markdown('<div class="sub-title">2. 风格偏好与输出要求</div>', unsafe_allow_html=True)
all_prefs = ["寓意好", "音韵美", "合五行", "合数理", "有典故", "有个性", "少歧义", "求时尚", "中性别", "地域性", "中英文"]
selected_prefs = st.multiselect("风格偏好 (可多选)", all_prefs, default=["寓意好", "合五行", "合数理", "音韵美", "有典故"])

c3, c4 = st.columns(2)
with c3: name_length = st.text_input("名字字数要求", value="3字")
with c4: name_count = st.number_input("生成方案个数", min_value=1, max_value=10, value=5)
other_req = st.text_area("其他补充要求 (如特定的字辈、避免的字等)")


# ================= 核心推演逻辑 =================
def get_bazi_info():
    gender_str = "乾造" if "男" in gender else "坤造"
    bazi_label = "法人八字" if mode == "为公司起名" else "命主八字"
    
    if not HAS_LUNAR:
        return f"公历时间：{cal_type} {year}年{month}月{day}日\n{bazi_label}：缺 (未安装排盘库) ({gender_str})"
    
    try:
        h = int(hour.split(" ")[0]) if hour != "不清楚时辰" else 12
        if cal_type == "公历":
            solar = Solar.fromYmdHms(int(year), int(month), int(day), h, 0, 0)
            lunar = solar.getLunar()
        else:
            lunar = Lunar.fromYmdHms(int(year), int(month), int(day), h, 0, 0)
            solar = lunar.getSolar()
            
        ba_zi = lunar.getEightChar()
        return f"出生公历：{solar.getYear()}年{solar.getMonth()}月{solar.getDay()}日 {h}时\n出生农历：{lunar.getYear()}年{lunar.getMonthInChinese()}月{lunar.getDayInChinese()} {h}时\n{bazi_label}：{ba_zi.getYear()} {ba_zi.getMonth()} {ba_zi.getDay()} {ba_zi.getTime()}（{gender_str}）"
    except Exception:
        return f"日期解析异常 ({gender_str})"

def generate_word_doc(content, bazi_info):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Microsoft YaHei'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    style.font.size = Pt(11)

    fav_element, taboo_element = "火", "水"
    color_match = re.search(r'【五行色彩：喜=(.)，忌=(.)】', content)
    if color_match:
        fav_element, taboo_element = color_match.group(1), color_match.group(2)
        content = content.replace(color_match.group(0), '').strip()

    color_map = {'木': RGBColor(0, 128, 0), '火': RGBColor(255, 0, 0), '土': RGBColor(255, 255, 0), '金': RGBColor(255, 255, 255), '水': RGBColor(0, 0, 0)}
    hex_map = {'木': 'E8F5E9', '火': 'FFEBEE', '土': 'FFF9C4', '金': 'F5F5F5', '水': 'E0E0E0'}

    title_text = "个人专属起名策划方案" if mode == "为人起名" else "公司品牌起名策划方案"
    title_para = doc.add_paragraph()
    title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title_para.paragraph_format.space_before = Pt(28)
    title_para.paragraph_format.space_after = Pt(5)
    title_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    pPr = title_para._element.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_map.get(taboo_element, 'FFFFFF'))
    pPr.append(shd)

    title_run = title_para.add_run(title_text)
    title_run.font.name = 'SimHei'
    title_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimHei')
    title_run.font.size = Pt(24)
    title_run.font.bold = True
    title_run.font.color.rgb = color_map.get(fav_element, RGBColor(0, 0, 0))

    time_para = doc.add_paragraph(f"测算时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    time_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    time_para.runs[0].font.size = Pt(10)
    time_para.runs[0].font.color.rgb = RGBColor(128, 128, 128)

    for line in content.split('\n'):
        line = line.strip()
        if not line: continue
        p = doc.add_paragraph()
        if "第一部分" in line or "第二部分" in line:
            run = p.add_run(line)
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 102, 204)
            if "第一部分" in line:
                doc.add_paragraph("")
                bazi_para = doc.add_paragraph(bazi_info)
                bazi_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                run_bazi = bazi_para.runs[0]
                run_bazi.font.size = Pt(12)
                run_bazi.font.bold = True
                run_bazi.font.color.rgb = RGBColor(139, 69, 19)
                doc.add_paragraph("")
        elif line.startswith("【") and "方案" in line:
            run = p.add_run(line)
            run.font.size = Pt(14)
            run.font.bold = True
            run.font.color.rgb = RGBColor(204, 102, 0)
        elif any(line.startswith(f"{i}.") for i in range(1, 7)):
            parts = line.split("：", 1)
            if len(parts) == 2:
                p.add_run(parts[0] + "：").font.bold = True
                p.runs[0].font.color.rgb = RGBColor(34, 139, 34)
                p.add_run(parts[1])
            else:
                p.add_run(line)
        else:
            p.add_run(line)

    disclaimer = doc.add_paragraph("\n[ 注：本起名方案仅供参考 ]")
    disclaimer.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    disclaimer.runs[0].font.size = Pt(10)
    disclaimer.runs[0].font.color.rgb = RGBColor(128, 128, 128)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio

# 生成按钮
st.divider()
if st.button("🚀 连接 AI 开始推演方案", type="primary", use_container_width=True):
    if not st.session_state.api_key:
        st.warning("⚠️ 管理员尚未配置 API 密钥，请先在左侧后台管理中心进行设置！")
    else:
        with st.spinner("算力运转中，正在为您结合八字与国学全力推演，请稍候..."):
            try:
                bazi_info = get_bazi_info()
                
                # 构造 Prompt
                prompt = "你是一个精通中国传统文化、周易八卦、五行生克、三才五格数理、诗词歌赋的资深国学起名大师。\n\n"
                prompt += f"【背景信息】\n- 性别：{gender}\n- 出生时间与八字：\n{bazi_info}\n（请在第一部分结合性别给出详细的五行喜忌分析）\n"

                num_match = re.search(r'\d+', name_length)
                n_total = int(num_match.group()) if num_match else 3
                
                if mode == "为人起名":
                    prompt += f"- 起名类型：个人起名\n- 姓氏：{surname}\n"
                    if birth_place: prompt += f"- 出生地：{birth_place}\n"
                    surname_len = len(surname) if surname else 1
                    n_names = max(1, n_total - surname_len)
                    mask_chars = "〇" * n_names
                    template_mask = f"{surname}{mask_chars}"
                    prompt += f"\n【⚠️ 核心铁律】\n1. 总汉字数必须绝对等于 {n_total} 个汉字！\n2. 强制填空：名字必须完全填入模板【 {template_mask} 】。\n"
                else:
                    prompt += f"- 起名类型：公司/品牌起名\n- 行业及经营范围：{comp_industry}\n"
                    if comp_prefix or comp_suffix: prompt += f"- 结构：{comp_prefix} + [核心名字] + {comp_suffix}\n"
                    prompt += f"\n【⚠️ 核心铁律】\n1. 核心商号必须绝对等于 {n_total} 个汉字！\n2. 强制填空：必须完全填入模板【 {'〇' * n_total} 】。\n"

                prompt += f"【偏好要求】\n- 勾选要求：{', '.join(selected_prefs)}\n- 补充要求：{other_req}\n"
                prompt += "\n【五行颜色底层指令提取】\n请在整篇回答的最开头（第一行），原封不动输出：\n【五行色彩：喜=X，忌=Y】（X和Y替换为单字，勿带其他文字）。\n"
                prompt += f"\n【输出任务】\n生成 {name_count} 个方案。严格按以下结构：\n第一部分：五行喜忌分析\n第二部分：起名方案\n【方案一：XXX】\n1. 拼音：\n2. 五行：\n3. 数理：\n4. 音韵：\n5. 典故：\n6. 寓意：\n"

                # 调用 API
                client = OpenAI(api_key=st.session_state.api_key, base_url=st.session_state.api_base)
                response = client.chat.completions.create(
                    model=st.session_state.model_name,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.7
                )
                
                ai_content = response.choices[0].message.content
                
                # 页面展示
                st.success("✅ 推演成功！")
                st.session_state.usage_count += 1  # 增加使用次数统计
                
                st.markdown(ai_content.replace(re.search(r'【五行色彩：喜=.*，忌=.*】', ai_content).group(0), '') if re.search(r'【五行色彩：喜=.*，忌=.*】', ai_content) else ai_content)

                # 生成 Word 内存文件并提供下载
                word_file = generate_word_doc(ai_content, bazi_info)
                filename = f"起名策划方案_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                
                st.download_button(
                    label="📥 保存为 Word 策划方案文档",
                    data=word_file.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"❌ 生成失败，请联系管理员检查后台配置或网络。详细报错：{e}")
