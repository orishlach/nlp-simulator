import streamlit as st
import json
from protocol_parser import process_document
import tempfile
import os
from PIL import Image
import pandas as pd

# Set page configuration with proper RTL support
st.set_page_config(
    page_title="🎤 מנתח פרוטוקולים של הכנסת",
    page_icon="🎤",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better Hebrew support and styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;700&display=swap');
    
    * {
        font-family: 'Heebo', sans-serif;
    }
    
    .main {
        direction: rtl;
        text-align: right;
    }
    
    h1, h2, h3, h4 {
        color: #003366;
        font-weight: 700;
    }
    
    .stButton button {
        background-color: #003366;
        color: white;
        border-radius: 5px;
        border: none;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    
    .stButton button:hover {
        background-color: #004d99;
    }
    
    .upload-section {
        background-color: #f8f9fa;
        padding: 2rem;
        border-radius: 10px;
        border: 1px solid #e9ecef;
        margin-bottom: 2rem;
    }
    
    .stSelectbox div[data-baseweb="select"] > div {
        direction: rtl;
        text-align: right;
    }
    
    .dataframe {
        direction: rtl;
        text-align: right;
    }
    
    footer {
        visibility: hidden;
    }
    
    .blue-header {
        background-color: #e6f2ff;
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 1rem;
        border-right: 4px solid #003366;
    }
    
    .metric-container {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 1rem;
        text-align: center;
        border: 1px solid #e9ecef;
    }
    
    .sentence-card {
        background-color: white;
        padding: 1rem;
        border-radius: 8px;
        border-right: 3px solid #007bff;
        margin-bottom: 1rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    }
    
    .speaker-tag {
        background-color: #e6f2ff;
        padding: 0.3rem 0.6rem;
        border-radius: 15px;
        font-weight: 500;
        color: #003366;
        display: inline-block;
        margin-bottom: 0.5rem;
    }
    
    .feature-card {
        background-color: white;
        padding: 1.5rem;
        border-radius: 8px;
        height: 100%;
        border: 1px solid #eee;
    }
    
    .welcome-container {
        text-align: center;
        padding: 3rem 1rem;
        background-color: #f8f9fa;
        border-radius: 10px;
        margin: 2rem 0;
    }
    
    .feature-box {
        background-color: #e6f2ff;
        padding: 1.5rem;
        border-radius: 8px;
    }
    
    .feature-grid {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 1rem;
        max-width: 800px;
        margin: 0 auto;
    }
    
    .divider {
        height: 1px;
        background-color: #eee;
        margin: 2rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Helper functions
def create_sidebar():
    with st.sidebar:
        try:
            # Try to load the logo - use a default header if it fails
            st.image("logo.png", width=100)
        except:
            st.markdown("<h3>🎤 מנתח פרוטוקולים</h3>", unsafe_allow_html=True)
            
        st.title("מנתח פרוטוקולים של הכנסת")
        st.markdown("---")
        st.subheader("אודות")
        st.write("""
        כלי זה מאפשר לכם לנתח פרוטוקולים של הכנסת באמצעות עיבוד שפה טבעית.
        העלו קובץ Word של פרוטוקול וקבלו ניתוח מפורט.
        """)
        st.markdown("---")
        st.subheader("מדריך שימוש")
        st.markdown("""
        1. העלו קובץ פרוטוקול בפורמט .docx
        2. המתינו לעיבוד הנתונים
        3. צפו בניתוח המפורט
        4. הורידו את התוצאות בפורמט JSONL
        """)
        st.markdown("---")
        st.caption("פותח במסגרת פרויקט NLP לעיבוד שפה טבעית")

def display_metrics(filtered_sentences):
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('<div class="metric-container">', unsafe_allow_html=True)
        st.metric("מספר משפטים", len(filtered_sentences))
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="metric-container">', unsafe_allow_html=True)
        st.metric("מספר דוברים", len(set(s['speaker_name'] for s in filtered_sentences)))
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="metric-container">', unsafe_allow_html=True)
        avg_sentence_length = sum(len(s['sentence_text'].split()) for s in filtered_sentences) / len(filtered_sentences)
        st.metric("אורך משפט ממוצע", f"{avg_sentence_length:.1f} מילים")
        st.markdown('</div>', unsafe_allow_html=True)

def visualize_data(filtered_sentences):
    # Create dataframe for visualization
    df = pd.DataFrame(filtered_sentences)
    
    # Speaker frequency chart
    st.markdown("""
    <div class="blue-header">
        <h3>ניתוח דוברים</h3>
        <p style="color: #666666;">התפלגות המשפטים לפי דובר</p>
    </div>
    """, unsafe_allow_html=True)
    
    speaker_counts = df['speaker_name'].value_counts().reset_index()
    speaker_counts.columns = ['דובר', 'מספר משפטים']
    
    if len(speaker_counts) > 10:
        top_speakers = speaker_counts.head(10)
        chart_data = top_speakers
        title = '10 הדוברים המובילים'
    else:
        chart_data = speaker_counts
        title = 'מספר משפטים לפי דובר'
    
    # Using Streamlit's native bar chart
    st.subheader(title)
    st.bar_chart(chart_data.set_index('דובר'))
    
    # Sentence length distribution
    df['sentence_length'] = df['sentence_text'].apply(lambda x: len(x.split()))
    
    st.markdown("""
    <div class="blue-header">
        <h3>אורך משפטים</h3>
        <p style="color: #666666;">התפלגות אורך המשפטים</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.subheader('התפלגות אורך המשפטים (מספר מילים)')
    st.bar_chart(df['sentence_length'].value_counts().sort_index())

def display_sentences(filtered_sentences):
    st.markdown("""
    <div class="blue-header">
        <h3>משפטים לדוגמה</h3>
        <p style="color: #666666;">הצגת משפטים מתוך הפרוטוקול</p>
    </div>
    """, unsafe_allow_html=True)
    
    for i, sentence in enumerate(filtered_sentences[:30]):
        st.markdown(f"""
        <div class="sentence-card">
            <div class="speaker-tag">{sentence['speaker_name']}</div>
            <p>{sentence['sentence_text']}</p>
            <small style="color: #6c757d;">פרוטוקול: {sentence['protocol_name']}</small>
        </div>
        """, unsafe_allow_html=True)

# Main application
def main():
    create_sidebar()
    
    # Main content
    st.title("🎤 מנתח פרוטוקולים של הכנסת")
    st.markdown("העלו קובצי פרוטוקולים של הכנסת בפורמט Word וקבלו ניתוח מבוסס NLP")
    
    # File upload section
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "בחרו קובץ או קבצי פרוטוקול בפורמט .docx",
        type="docx",
        accept_multiple_files=True,
        key="protocol_uploader"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_files:
        with st.spinner("מעבד את הקבצים, אנא המתן..."):
            all_sentences = []
            for uploaded_file in uploaded_files:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_file_path = tmp_file.name
                
                protocol_name = uploaded_file.name
                knesset_number, protocol_type = 24, 'מליאה'  # ניתן להרחיב
                
                try:
                    protocol = process_document(tmp_file_path, protocol_name, knesset_number, protocol_type)
                    if protocol:
                        for sentence in protocol.sentences:
                            all_sentences.append({
                                'protocol_name': protocol.protocol_name,
                                'knesset_number': protocol.knesset_number,
                                'protocol_type': protocol.protocol_type,
                                'protocol_number': protocol.protocol_number,
                                'speaker_name': sentence.speaker_name,
                                'sentence_text': sentence.sentence_text,
                            })
                except Exception as e:
                    st.error(f"שגיאה בעיבוד הקובץ {protocol_name}: {str(e)}")
                
                os.unlink(tmp_file_path)
            
            if all_sentences:
                # Success message
                st.success(f"✅ הוצאו בהצלחה {len(all_sentences)} משפטים מהפרוטוקולים!")
                
                # Display metrics
                display_metrics(all_sentences)
                
                # Create a select box for navigation
                page = st.selectbox(
                    "בחר תצוגה:",
                    ["📊 ויזואליזציה", "📝 משפטים", "⚙️ סינון מתקדם"]
                )
                
                st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                
                if page == "📊 ויזואליזציה":
                    visualize_data(all_sentences)
                
                elif page == "📝 משפטים":
                    display_sentences(all_sentences)
                
                elif page == "⚙️ סינון מתקדם":
                    st.subheader("סננו את התוצאות")
                    
                    # Filters
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        unique_speakers = sorted(set(s['speaker_name'] for s in all_sentences))
                        speaker_filter = st.selectbox(
                            "סננו לפי דובר:",
                            ["הכל"] + unique_speakers
                        )
                    
                    with col2:
                        unique_protocols = sorted(set(s['protocol_name'] for s in all_sentences))
                        protocol_filter = st.selectbox(
                            "סננו לפי פרוטוקול:",
                            ["הכל"] + unique_protocols
                        )
                    
                    # Text search
                    text_search = st.text_input("חיפוש טקסט:")
                    
                    # Apply filters
                    filtered_sentences = all_sentences
                    
                    if speaker_filter != "הכל":
                        filtered_sentences = [s for s in filtered_sentences if s['speaker_name'] == speaker_filter]
                    
                    if protocol_filter != "הכל":
                        filtered_sentences = [s for s in filtered_sentences if s['protocol_name'] == protocol_filter]
                    
                    if text_search:
                        filtered_sentences = [s for s in filtered_sentences if text_search in s['sentence_text']]
                    
                    # Display filtered results
                    st.write(f"נמצאו {len(filtered_sentences)} תוצאות")
                    if filtered_sentences:
                        df_filtered = pd.DataFrame(filtered_sentences)
                        st.dataframe(df_filtered[['speaker_name', 'sentence_text', 'protocol_name']])
                
                # Download section
                st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                st.subheader("הורדת הנתונים")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    jsonl_data = "\n".join(json.dumps(line, ensure_ascii=False) for line in all_sentences)
                    st.download_button(
                        "📥 הורד בפורמט JSONL",
                        jsonl_data,
                        file_name="protocol_sentences.jsonl",
                        mime="application/jsonl",
                        key="download_jsonl"
                    )
                
                with col2:
                    csv_data = pd.DataFrame(all_sentences).to_csv(index=False)
                    st.download_button(
                        "📥 הורד בפורמט CSV",
                        csv_data,
                        file_name="protocol_sentences.csv",
                        mime="text/csv",
                        key="download_csv"
                    )
        
    else:
        # Display welcome message when no files uploaded
        st.markdown("""
        <div class="welcome-container">
            <h2 style="margin-bottom: 1rem; color: #003366;">ברוכים הבאים למנתח פרוטוקולים של הכנסת</h2>
            <p style="font-size: 1.1rem; margin-bottom: 2rem;">העלו קובצי פרוטוקול בפורמט .docx כדי להתחיל בניתוח</p>
            
            <div class="feature-grid">
                <div class="feature-box">
                    <h3 style="color: #003366; margin-bottom: 0.5rem;">ניתוח טקסט</h3>
                    <p>עיבוד שפה טבעית לפרוטוקולים של הכנסת</p>
                </div>
                <div class="feature-box">
                    <h3 style="color: #003366; margin-bottom: 0.5rem;">ויזואליזציה</h3>
                    <p>תצוגות גרפיות להבנה מהירה של הנתונים</p>
                </div>
                <div class="feature-box">
                    <h3 style="color: #003366; margin-bottom: 0.5rem;">ייצוא נתונים</h3>
                    <p>הורדת תוצאות הניתוח בפורמטים שונים</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Features explanation
        st.subheader("תכונות המערכת:")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            <div class="feature-card">
                <h4 style="color: #003366; margin-bottom: 1rem;">ניתוח פרוטוקולים</h4>
                <ul>
                    <li>זיהוי אוטומטי של דוברים</li>
                    <li>חילוץ משפטים מתוך הטקסט</li>
                    <li>ניתוח סטטיסטי של הדיונים</li>
                    <li>סינון מתקדם לפי דובר ותוכן</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
            
        with col2:
            st.markdown("""
            <div class="feature-card">
                <h4 style="color: #003366; margin-bottom: 1rem;">יתרונות המערכת</h4>
                <ul>
                    <li>עיבוד מהיר של קבצי Word</li>
                    <li>ממשק משתמש נוח ואינטואיטיבי</li>
                    <li>אפשרויות ייצוא מתקדמות</li>
                    <li>תמיכה בעברית מלאה</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

# Run the application
if __name__ == "__main__":
    main()
