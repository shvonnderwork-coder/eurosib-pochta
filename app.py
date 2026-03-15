import streamlit as st
import streamlit.components.v1 as components
from docxtpl import DocxTemplate
import requests
import io
import re
import os
import base64

# --- НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Почтовые документы", page_icon="🖨️", layout="centered")

# === ВАШИ ДАННЫЕ ===
DADATA_API_KEY = "7e2ff734d7cec9715ad889d4efdd7c32c26e78dc" 
SENDER_NAME = 'АО "ЕВРОСИБ СПБ-ТРАНСПОРТНЫЕ СИСТЕМЫ"'

num_to_words = {
    1: "Один", 2: "Два", 3: "Три", 4: "Четыре", 5: "Пять",
    6: "Шесть", 7: "Семь", 8: "Восемь", 9: "Девять", 10: "Десять",
    11: "Одиннадцать", 12: "Двенадцать", 13: "Тринадцать", 14: "Четырнадцать", 15: "Пятнадцать"
}

def get_company_info(inn, api_key):
    url = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"
    headers = {"Content-Type": "application/json", "Authorization": f"Token {api_key}"}
    try:
        response = requests.post(url, headers=headers, json={"query": inn}, timeout=10)
        if response.status_code == 200 and response.json().get("suggestions"):
            data = response.json()["suggestions"][0]["data"]
            short_name = data["name"].get("short_with_opf") or data["name"]["full_with_opf"]
            return {
                "RECEIVER_NAME": data["name"]["full_with_opf"],
                "RECEIVER_SHORT_NAME": short_name,
                "RECEIVER_ADDRESS": data["address"]["value"],
                "RECEIVER_INDEX": data.get("address", {}).get("data", {}).get("postal_code") or ""
            }
    except:
        return None
    return None

# --- ИНТЕРФЕЙС ПРИЛОЖЕНИЯ ---
st.title("🖨️ Генератор почтовых документов")
st.markdown(f"**Отправитель:** {SENDER_NAME}")

# Поле для ИНН
inn_input = st.text_input("ИНН получателя:", placeholder="Например: 7707083893")

# Поля для документов
st.markdown("### Список вложений (до 10 строк):")
items_inputs = []
for i in range(1, 11):
    doc = st.text_input(f"Документ {i}", key=f"doc_{i}")
    if doc.strip():
        items_inputs.append(doc.strip())

# Кнопка запуска
if st.button("Сгенерировать документы", type="primary"):
    if not inn_input:
        st.warning("Пожалуйста, введите ИНН.")
    elif not os.path.exists("Опись вложения.docx") or not os.path.exists("Кому куда.docx"):
        st.error("❌ Ошибка: Шаблоны Word не найдены в папке приложения!")
    else:
        with st.spinner("Запрашиваем данные из DaData и формируем файлы..."):
            c_data = get_company_info(inn_input.strip(), DADATA_API_KEY)
            
            if not c_data:
                st.error("❌ Компания по этому ИНН не найдена.")
            else:
                data = {"SENDER": SENDER_NAME}
                for i in range(1, 16):
                    if i <= len(items_inputs):
                        data[f"N_{i}"] = str(i)
                        data[f"ITEM_{i}"] = items_inputs[i-1]
                        data[f"COUNT_{i}"] = "1"
                        data[f"VALUE_{i}"] = "1"
                    else:
                        data[f"N_{i}"] = data[f"ITEM_{i}"] = data[f"COUNT_{i}"] = data[f"VALUE_{i}"] = ""
                
                count = len(items_inputs)
                word = num_to_words.get(count, str(count))
                data["TOTAL_COUNT"] = str(count)
                data["TOTAL_VALUE"] = f"{count} ({word}) руб. 00 коп." if count > 0 else "0 (Ноль) руб. 00 коп."

                # Рендерим файлы
                doc_o = DocxTemplate("Опись вложения.docx")
                doc_o.render(data)
                o_buf = io.BytesIO()
                doc_o.save(o_buf)
                
                doc_k = DocxTemplate("Кому куда.docx")
                doc_k.render(c_data)
                k_buf = io.BytesIO()
                doc_k.save(k_buf)
                
                safe_name = re.sub(r'[\\/*?:"<>|]', "", c_data['RECEIVER_SHORT_NAME']).strip()
                
                st.success(f"✅ Документы для **{safe_name}** успешно созданы! Загрузка начнется автоматически.")
                
                # --- АВТОМАТИЧЕСКОЕ СКАЧИВАНИЕ ЧЕРЕЗ JAVASCRIPT ---
                b64_o = base64.b64encode(o_buf.getvalue()).decode()
                b64_k = base64.b64encode(k_buf.getvalue()).decode()
                
                js_code = f"""
                <script>
                    function downloadFile(filename, data) {{
                        var element = document.createElement('a');
                        element.setAttribute('href', 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + data);
                        element.setAttribute('download', filename);
                        element.style.display = 'none';
                        document.body.appendChild(element);
                        element.click();
                        document.body.removeChild(element);
                    }}
                    
                    // Скачиваем первый файл
                    downloadFile("Опись_{safe_name}.docx", "{b64_o}");
                    
                    // Ждем полсекунды и скачиваем второй
                    setTimeout(function() {{
                        downloadFile("Конверт_{safe_name}.docx", "{b64_k}");
                    }}, 600);
                </script>
                """
                components.html(js_code, height=0)
                
                # Запасные кнопки на случай, если браузер заблокирует всплывающие окна
                st.markdown("*Если скачивание не началось автоматически, нажмите на кнопки ниже:*")
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(label="📄 Скачать Опись", data=o_buf.getvalue(), file_name=f"Опись_{safe_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                with col2:
                    st.download_button(label="✉️ Скачать Конверт", data=k_buf.getvalue(), file_name=f"Конверт_{safe_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
