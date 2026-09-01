import streamlit as st
import os
from datetime import datetime
from docx import Document
from openai import OpenAI

try:
    from lunar_python import Solar, Lunar
    HAS_LUNAR = True
except ImportError:
    HAS_LUNAR = False

st.set_page_config(page_title="国学传统文化 AI 起名系统", layout="centered")

st.title("🌟 国学传统文化 AI 起名系统")

# 1. 侧边栏：AI 模型配置
st.sidebar.header("1. AI 模型配置")
api_base = st.sidebar.text_input("接口地址 (Base URL)", value="https://api.openai.com/v1")
api_key = st.sidebar.text_input("API 密钥 (API Key)", type="password")
model_name = st.sidebar.text_input("模型名称", value="gpt-4o")

# 2. 主界面：模式选择
mode = st.radio("选择起名类型", ["为人起名", "为公司起名"])

col1, col2 = st.columns(2)
with col1:
    cal_type = st.selectbox("历法", ["公历", "农历"])
    year = st.number_input("年", min_value=1920, max_value=2040, value=1990)
    month = st.number_input("月", min_value=1, max_value=12, value=1)
with col2:
    day = st.number_input("日", min_value=1, max_value=31, value=1)
    hour = st.slider("时辰 (24小时制)", 0, 23, 12)
    gender = st.radio("性别/属性", ["男 (乾造)", "女 (坤造)"], horizontal=True)

if mode == "为人起名":
    surname = st.text_input("姓氏")
    birth_place = st.text_input("出生地 (选填)")
else:
    comp_prefix = st.text_input("公司前缀 (如XX市)")
    comp_suffix = st.text_input("公司后缀 (如科技有限公司)")
    comp_industry = st.text_input("行业及业务范围")

name_length = st.text_input("名字字数要求", value="3字")
name_count = st.number_input("生成方案个数", min_value=1, max_value=10, value=3)
other_req = st.text_area("其他补充要求")

# 生成逻辑
if st.button("🚀 开始 AI 推演生成方案", type="primary"):
    if not api_key:
        st.error("请先在左侧边栏填写 API 密钥！")
    else:
        with st.spinner("AI 正在结合八字与国学全力推演中，请稍候..."):
            try:
                # 八字计算
                gender_str = "乾造" if "男" in gender else "坤造"
                bazi_info = ""
                if HAS_LUNAR:
                    if cal_type == "公历":
                        solar = Solar.fromYmdHms(int(year), int(month), int(day), int(hour), 0, 0)
                        lunar = solar.getLunar()
                    else:
                        lunar = Lunar.fromYmdHms(int(year), int(month), int(day), int(hour), 0, 0)
                        solar = lunar.getSolar()
                    ba_zi = lunar.getEightChar()
                    bazi_info = f"公历：{solar.toYmdHms()} / 农历：{lunar.toFullString()}\n八字：{ba_zi.getYear()} {ba_zi.getMonth()} {ba_zi.getDay()} {ba_zi.getTime()}（{gender_str}）"
                else:
                    bazi_info = f"时间：{year}年{month}月{day}日 {hour}时 ({gender_str})"

                # 构造 Prompt 
                prompt = f"你是一个精通国学起名的专家。请根据以下信息生成 {name_count} 个名字方案，字数要求：{name_length}。\n"
                prompt += f"八字信息：\n{bazi_info}\n"
                if mode == "person":
                    prompt += f"个人起名，姓氏：{surname}\n"
                else:
                    prompt += f"公司起名，全称结构：{comp_prefix} + [核心] + {comp_suffix}，行业：{comp_industry}\n"
                if other_req:
                    prompt += f"其他要求：{other_req}\n"
                
                prompt += "\n请严格按以下结构输出：\n第一部分：五行喜忌分析\n第二部分：起名方案（包含拼音、五行、数理、音韵、典故、寓意）"

                # 调用 OpenAI 接口
                client = OpenAI(api_key=api_key, base_url=api_base)
                response = client.chat.completions.create(
                    model=model_name,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.7
                )
                ai_content = response.choices[0].message.content

                st.success("🎉 起名方案生成成功！")
                st.markdown(ai_content)

                # 生成 Word 下载
                doc = Document()
                doc.add_heading("国学起名策划方案", 0)
                doc.add_paragraph(bazi_info)
                doc.add_paragraph(ai_content)
                
                filename = f"起名方案_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                doc.save(filename)

                with open(filename, "rb") as file:
                    st.download_button(
                        label="📥 下载 Word 策划方案文档",
                        data=file,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"发生错误: {e}")