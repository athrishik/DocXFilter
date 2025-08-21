#!/usr/bin/env python
# coding: utf-8

"""
DocXFilter v2.0 - Advanced Document Search & Analytics
Copyright 2025 Hrishik Kunduru. All rights reserved.

Professional multi-pattern search and analytics tool for DocXScan Excel outputs.
"""

import streamlit as st
import pandas as pd
import json
import re
import io
from datetime import datetime
from collections import defaultdict
import plotly.express as px
from typing import Dict, List

# Configure Streamlit page
st.set_page_config(
    page_title="DocXFilter",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #2563EB, #10B981);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .token-highlight {
        background-color: #ffd700;
        padding: 2px 4px;
        border-radius: 3px;
        font-weight: bold;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { height: 50px; padding-left: 20px; padding-right: 20px; }
</style>
""", unsafe_allow_html=True)

class DocumentSearchEngine:
    """Document search engine with caching"""
    def __init__(self):
        self.df = None
        self.token_map = {}
        self.discovered_tokens = {}
        self._search_cache = {}

    def load_data(self, df: pd.DataFrame) -> bool:
        """Load Excel data"""
        try:
            if 'File Name' not in df.columns or 'Full Contents' not in df.columns:
                return False
            self.df = df.copy()
            self.df['Full Contents'] = self.df['Full Contents'].fillna('').astype(str)
            self.df['File Name'] = self.df['File Name'].astype(str)
            self._search_cache.clear()
            return True
        except Exception:
            return False

    def load_token_definitions(self, token_json: dict):
        """Load token definitions from JSON"""
        self.token_map = token_json

    @st.cache_data
    def discover_tokens(_self) -> Dict[str, Dict]:
        """Discover all tokens with caching"""
        if _self.df is None: return {}
        patterns = [
            r'<<[^>]+>>',
            r'<<[^>]+\.',
            r'\{[A-Z_][^}]*\}',
            r'\[[A-Z_][^\]]*\]',
            r'\[\[[A-Z_][^\]]*',
            r'<[a-z]+>',
        ]
        compiled_patterns = [re.compile(p, re.IGNORECASE) for p in patterns]
        all_tokens = defaultdict(lambda: {'count': 0, 'documents': set()})
        for _, row in _self.df.iterrows():
            content = str(row['Full Contents'])
            file_name = row['File Name']
            for pattern in compiled_patterns:
                for match in pattern.finditer(content):
                    token = match.group()
                    all_tokens[token]['count'] += 1
                    all_tokens[token]['documents'].add(file_name)
        result = {token: {
            'count': data['count'],
            'doc_count': len(data['documents']),
            'documents': list(data['documents'])[:20]
        } for token, data in all_tokens.items()}
        _self.discovered_tokens = result
        return result

    def search_multi(self, search_terms: List[str], mode: str) -> pd.DataFrame:
        """Search for documents containing all/any search terms"""
        if self.df is None or not search_terms:
            return pd.DataFrame()
        result_df = self.df.copy()
        if mode == "AND":
            for term in search_terms:
                mask = result_df['Full Contents'].str.contains(term, case=False, na=False, regex=False)
                result_df = result_df[mask]
        else: # OR
            all_masks = [
                self.df['Full Contents'].str.contains(term, case=False, na=False, regex=False)
                for term in search_terms
            ]
            if not all_masks:
                return pd.DataFrame()
            combined_mask = all_masks[0]
            for mask in all_masks[1:]:
                combined_mask = combined_mask | mask
            result_df = self.df[combined_mask]
        return result_df

    def get_contexts(self, search_term: str, doc_name: str, context_length: int = 100) -> List[str]:
        """Context around search term in document"""
        doc_row = self.df[self.df['File Name'] == doc_name]
        if doc_row.empty:
            return []
        content = str(doc_row['Full Contents'].iloc[0])
        contexts = []
        search_lower = search_term.lower()
        content_lower = content.lower()
        start_pos = 0
        while len(contexts) < 3:
            pos = content_lower.find(search_lower, start_pos)
            if pos == -1: break
            context_start = max(0, pos - context_length)
            context_end = min(len(content), pos + len(search_term) + context_length)
            context = content[context_start:context_end].replace('\n', ' ').strip()
            if context and context not in contexts:
                contexts.append(context)
            start_pos = pos + 1
        return contexts

# --- SESSION STATE INIT ---
if 'search_engine' not in st.session_state:
    st.session_state.search_engine = DocumentSearchEngine()
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'current_results' not in st.session_state:
    st.session_state.current_results = pd.DataFrame()
if 'search_terms' not in st.session_state:
    st.session_state.search_terms = []
if 'search_mode' not in st.session_state:
    st.session_state.search_mode = "AND"
if 'current_search_key' not in st.session_state:
    st.session_state.current_search_key = ""
if 'input_counter' not in st.session_state:
    st.session_state.input_counter = 0

# --- APP MAIN ---
def main():
    st.markdown("""
    <div class="main-header">
        <h1>🔍 DocXFilter</h1>
        <p>Advanced multi-pattern search and analytics for DocXScan outputs</p>
    </div>
    """, unsafe_allow_html=True)
    with st.sidebar:
        st.header("📁 Data Import")
        uploaded_excel = st.file_uploader(
            "DocXScan Excel File", type=['xlsx', 'xls'], key="excel_uploader")
        uploaded_tokens = st.file_uploader(
            "Token Definitions (Optional)", type=['json'], key="token_uploader")
        if uploaded_excel and not st.session_state.data_loaded:
            try:
                df = pd.read_excel(uploaded_excel, engine='openpyxl')
                if st.session_state.search_engine.load_data(df):
                    if uploaded_tokens:
                        token_json = json.load(uploaded_tokens)
                        st.session_state.search_engine.load_token_definitions(token_json)
                        st.success(f"✅ Loaded {len(token_json)} token definitions")
                    with st.spinner("Discovering patterns..."):
                        st.session_state.search_engine.discover_tokens()
                    st.session_state.data_loaded = True
                    st.success(f"✅ Loaded {len(df)} documents")
                    st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
        if st.session_state.data_loaded:
            st.markdown("---")
            if st.button("🔄 Reset & Load New Files"):
                st.session_state.data_loaded = False
                st.session_state.current_results = pd.DataFrame()
                st.session_state.search_terms = []
                st.session_state.current_search_key = ""
                st.session_state.search_engine = DocumentSearchEngine()
                st.rerun()
    # Main interface
    if st.session_state.data_loaded:
        show_main_interface()
    else:
        show_welcome_screen()

def show_welcome_screen():
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown("""
        ### 👋 Welcome to DocXFilter
        
        **Advanced Document Search & Analytics Tool:**
        - 🔍 Multi-pattern search with AND/OR logic
        - 📊 Document-level analytics and insights
        - 💡 Pattern discovery
        - 📈 Data quality assessment
        - 🎯 Export Excel reports
        
        **Quick Start:**
        1. Upload your DocXScan Excel file in the sidebar
        2. Optionally upload JSON token definitions
        3. Add multiple search terms and choose AND/OR mode
        4. Explore analytics for deep insights
        5. Export reports
        """)

def show_main_interface():
    engine = st.session_state.search_engine
    total_docs = len(engine.df) if engine.df is not None else 0
    total_tokens = len(engine.discovered_tokens)
    results_count = len(st.session_state.current_results)
    col1, col2, col3 = st.columns(3)
    with col1: st.metric("📄 Documents", total_docs)
    with col2: st.metric("🔑 Patterns Found", total_tokens)
    with col3: st.metric("📋 Search Results", results_count)
    st.markdown("---")
    tab1, tab2, tab3 = st.tabs(["🔍 Search", "📊 Analytics", "📤 Export"])
    with tab1: show_search_interface()
    with tab2: show_analytics_interface()
    with tab3: show_export_interface()

def show_search_interface():
    st.subheader("🔍 Multi-Pattern Search")
    engine = st.session_state.search_engine
    token_options = []

    # Merge JSON tokens + discovered tokens for Quick Pattern Add
    if engine.token_map:
        for token, desc in engine.token_map.items():
            token_options.append(f"{token} - {desc}")
    used_tokens = set(engine.token_map.keys())
    for token in sorted(engine.discovered_tokens.keys()):
        if token not in used_tokens:
            doc_count = engine.discovered_tokens[token]['doc_count']
            token_options.append(f"{token} ({doc_count} docs)")

    # --- Term input ---
    st.markdown("#### ➕ Add Search Terms")
    col1, col2, col3 = st.columns([3,1,1])
    with col1:
        input_key = f"new_search_term_{st.session_state.input_counter}"
        new_term = st.text_input(
            "Enter text to search:", placeholder="Examples: <<lmerge, {CLIENT_NAME}, contract, SIGNATURE...", key=input_key)
    with col2:
        if st.button("➕ Add Term", type="primary", use_container_width=True):
            if new_term and new_term.strip():
                clean_term = new_term.strip()
                if clean_term not in st.session_state.search_terms:
                    st.session_state.search_terms.append(clean_term)
                    st.session_state.input_counter += 1
                    st.rerun()
                else:
                    st.warning(f"'{clean_term}' is already in your search terms!")
    with col3:
        if st.button("🧹 Clear All", use_container_width=True):
            st.session_state.search_terms = []
            st.session_state.current_results = pd.DataFrame()
            st.session_state.input_counter += 1
            st.session_state.current_search_key = ""
            st.rerun()

    if token_options:
        with st.expander("⚡ Quick Add from Discovered Patterns"):
            col1, col2 = st.columns([3,1])
            with col1:
                selected_pattern = st.selectbox(
                    "Choose a pattern:", [""] + token_options[:50], key="pattern_selector")
            with col2:
                st.write("") # spacer
                if st.button("➕ Add Pattern", use_container_width=True):
                    if selected_pattern:
                        token = selected_pattern.split(' - ')[0].split(' (')[0]
                        if token not in st.session_state.search_terms:
                            st.session_state.search_terms.append(token)
                            st.rerun()
                        else:
                            st.warning(f"'{token}' already in your search terms!")

    # --- Show current terms & search mode ---
    if st.session_state.search_terms:
        st.markdown("#### 🔍 Current Search Terms")
        col1, col2 = st.columns([2,1])
        # Search mode selection
        with col1:
            mode = st.radio(
                "Search Mode:", ["AND", "OR"], index=0 if st.session_state.search_mode == "AND" else 1,
                horizontal=True,
                help="AND: All terms must be present | OR: Any term present")
            st.session_state.search_mode = mode
        with col2:
            if st.button("🔍 Search Now", type="primary", use_container_width=True):
                perform_multi_search()

        # List terms with [x] remove button
        st.markdown("**Active Search Terms:**")
        for i, term in enumerate(st.session_state.search_terms):
            colT, colR = st.columns([6,1])
            with colT:
                st.markdown(f"🔍 `{term}`")
            with colR:
                if st.button("❌", key=f"remove_term_{i}", help=f"Remove '{term}'"):
                    st.session_state.search_terms.pop(i)
                    st.session_state.current_search_key = ""
                    st.rerun()

        # Run search automatically on change
        key = f"{mode}:{'|'.join(sorted(st.session_state.search_terms))}"
        if key != st.session_state.current_search_key and len(st.session_state.search_terms) > 0:
            perform_multi_search()

    else:
        st.info("💡 Add search terms above to start searching. You can add multiple patterns and search for documents containing all (AND) or any (OR) of them.")

    # --- Results Display ---
    if not st.session_state.current_results.empty:
        show_search_results()

def perform_multi_search():
    terms = st.session_state.search_terms
    mode = st.session_state.search_mode
    engine = st.session_state.search_engine
    result_df = engine.search_multi(terms, mode)
    st.session_state.current_results = result_df
    st.session_state.current_search_key = f"{mode}:{'|'.join(sorted(terms))}"
    if not result_df.empty:
        st.success(f"✅ Found {len(result_df)} documents containing {', '.join([f'`{t}`' for t in terms])}")
    else:
        st.warning(f"⚠️ No documents found for terms: {', '.join([f'`{t}`' for t in terms])}")

def show_search_results():
    results = st.session_state.current_results
    search_terms = st.session_state.search_terms
    search_mode = st.session_state.search_mode
    st.markdown("---")
    st.subheader(f"📋 Results ({len(results)} documents)")
    st.markdown(f"**Search:** {search_mode} search for: {', '.join([f'`{t}`' for t in search_terms])}")

    # Export button
    col1, col2 = st.columns([3,1])
    with col2:
        if st.button("📤 Export Results", use_container_width=True):
            export_data = create_enhanced_export(results, search_terms, search_mode)
            filename = f"multi_search_{search_mode.lower()}_{len(search_terms)}terms.xlsx"
            st.download_button(
                "💾 Download Excel", data=export_data, file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Table: file name, size, content length, match summary
    display_cols = ['File Name']
    for col in ['Size (KB)', 'Content Length (chars)']:
        if col in results.columns:
            display_cols.append(col)
    # Add match summary
    results_with_matches = results.copy()
    if not results_with_matches.empty:
        match_summaries = []
        for _, row in results_with_matches.iterrows():
            content = str(row['Full Contents']).lower()
            matches = []
            for term in search_terms:
                count = content.count(term.lower())
                if count > 0:
                    matches.append(f"{term}:{count}")
            match_summaries.append(" | ".join(matches))
        results_with_matches['Match Summary'] = match_summaries
        display_cols.append('Match Summary')
        st.dataframe(
            results_with_matches[display_cols],
            use_container_width=True,
            hide_index=True,
            height=300
        )

    # --- Document Preview ---
    if len(results) > 0:
        st.subheader("📖 Document Preview")
        selected_doc = st.selectbox(
            "Select document to preview:", results['File Name'].tolist(),
            help="Choose a document to see all matched terms in context")
        if selected_doc:
            show_enhanced_document_preview(selected_doc, search_terms)

def show_enhanced_document_preview(doc_name: str, search_terms: List[str]):
    engine = st.session_state.search_engine
    doc_row = engine.df[engine.df['File Name'] == doc_name]
    if doc_row.empty:
        st.error("Document not found"); return
    content = str(doc_row['Full Contents'].iloc[0])
    col1, col2 = st.columns(2)
    with col1: st.write(f"**📄 {doc_name}**")
    with col2:
        st.write("**🔍 Term Occurrences:**")
        for term in search_terms:
            count = content.lower().count(term.lower())
            st.write(f"• `{term}`: {count} times")
    st.write("**📋 Context Preview for Each Term:**")
    for i, term in enumerate(search_terms):
        with st.expander(f"🔍 Contexts for '{term}'", expanded=(i==0)):
            contexts = engine.get_contexts(term, doc_name)
            if contexts:
                # Highlight ALL search terms in context
                for j, context in enumerate(contexts, 1):
                    highlighted = context
                    colors = ['#ffd700', '#ffb3ba', '#bae1ff', '#baffc9', '#ffffba']
                    for k, highlight_term in enumerate(search_terms):
                        color = colors[k % len(colors)]
                        highlighted = re.sub(
                            re.escape(highlight_term),
                            f'<span style="background-color: {color}; padding: 2px 4px; border-radius: 3px; font-weight: bold;">{highlight_term}</span>',
                            highlighted, flags=re.IGNORECASE
                        )
                    st.markdown(
                        f'**Context {j}:** <div style="background: #f8f9fa; padding: 0.5rem; border-radius: 4px; margin: 0.5rem 0; border-left: 3px solid #2563EB;">{highlighted}</div>',
                        unsafe_allow_html=True
                    )
            else:
                st.write(f"No contexts found for '{term}' in this document.")
    st.markdown("#### 📝 Full Document Content (Matched Terms Highlighted)")
    colors = ['#ffd700', '#ffb3ba', '#bae1ff', '#baffc9', '#ffffba']

    highlighted_content = content
    # Replace longer terms first to avoid overlapping highlights
    for i, term in sorted(enumerate(search_terms), key=lambda x: -len(x[1])):
        color = colors[i % len(colors)]
        highlighted_content = re.sub(
            re.escape(term),
            f'<span style="background-color: {color}; padding:2px 4px; border-radius:3px; font-weight:bold;">{term}</span>',
            highlighted_content, flags=re.IGNORECASE
        )

    st.markdown(
        f"""
        <div style="
            background: #f8f9fa;
            padding: 1rem;
            border-radius: 8px;
            border: 1px solid #e9ecef;
            max-height: 600px;
            overflow-y: auto;
            white-space: pre-wrap;
            font-family: 'Courier New', monospace;
            font-size: 14px;
            line-height: 1.4;">
            {highlighted_content}
        </div>
        """,
        unsafe_allow_html=True
    )
@st.cache_data
def create_enhanced_export(results_df: pd.DataFrame, search_terms: List[str], search_mode: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df = results_df.copy()
        # Add match columns
        for term in search_terms:
            export_df[f"'{term}' Count"] = export_df['Full Contents'].str.lower().str.count(term.lower())
        export_df['Total Matches'] = sum([export_df[f"'{term}' Count"] for term in search_terms])
        export_cols = ['File Name', 'Size (KB)', 'Content Length (chars)'] + [f"'{term}' Count" for term in search_terms] + ['Total Matches']
        export_cols = [col for col in export_cols if col in export_df.columns]
        export_df[export_cols].to_excel(writer, sheet_name='Search Results', index=False)
        # Summary sheet
        summary_data = {
            'Search Mode': [search_mode],
            'Search Terms': [' | '.join(search_terms)],
            'Number of Terms': [len(search_terms)],
            'Results Found': [len(results_df)],
            'Export Date': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            'Generated By': ['DocXFilter v2.0']
        }
        for i, term in enumerate(search_terms, 1):
            summary_data[f'Term {i}'] = [term]
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Search Summary', index=False)
    output.seek(0)
    return output.read()

def show_analytics_interface():
    st.subheader("📊 Document Analytics Dashboard")
    engine = st.session_state.search_engine
    if engine.df is None or len(engine.df) == 0:
        st.info("No documents available for analytics."); return
    # Four analytics tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📄 Document Overview", 
        "📈 Content Analysis", 
        "🔍 Pattern Insights", 
        "📊 Export Reports"
    ])
    with tab1:
        show_document_overview(engine)
    with tab2:
        show_content_analysis(engine)
    with tab3:
        show_pattern_insights(engine)
    with tab4:
        show_export_reports(engine)

def show_document_overview(engine):
    st.markdown("### 📄 Document Collection Overview")
    df = engine.df
    # Key metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("📄 Total Documents", len(df))
    with col2: st.metric("💾 Total Size", f"{df.get('Size (KB)', pd.Series()).sum():,.1f} KB" if 'Size (KB)' in df.columns else "N/A")
    with col3:
        content_lengths = df['Full Contents'].str.len()
        st.metric("📝 Total Characters", f"{content_lengths.sum():,}")
    with col4:
        st.metric("📊 Avg Doc Length", f"{content_lengths.mean():,.0f} chars")
    # Distributions
    col1, col2 = st.columns(2)
    with col1:
        if 'Size (KB)' in df.columns and df['Size (KB)'].notna().any():
            fig = px.histogram(x=df['Size (KB)'].dropna(),
                nbins=25, title='📦 Document Size Distribution',
                labels={'x':'File Size (KB)','y':'#Documents'})
            fig.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No file size information available")
    with col2:
        fig = px.histogram(
            x=content_lengths,
            nbins=25,
            title='📝 Content Length Distribution',
            labels={'x': 'Content Length (characters)', 'y': 'Number of Documents'}
        )
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    # Document summary table
    st.markdown("### 📋 Document Statistics Summary")
    doc_analysis = []
    for _, row in df.iterrows():
        content = str(row['Full Contents'])
        word_count = len(content.split()) if content.strip() else 0
        line_count = content.count('\n') + 1 if content.strip() else 0
        angle_brackets = len(re.findall(r'<<[^>]*>>', content))
        curly_brackets = len(re.findall(r'\{[^}]*\}', content))
        square_brackets = len(re.findall(r'\[[^\]]*\]', content))
        total_patterns = angle_brackets + curly_brackets + square_brackets
        doc_analysis.append({
            'File Name': row['File Name'],
            'Size (KB)': row.get('Size (KB)', 0),
            'Content Length': len(content),
            'Word Count': word_count,
            'Line Count': line_count,
            'Pattern Count': total_patterns,
            'Content Density': word_count / max(len(content),1) * 100
        })
    doc_stats_df = pd.DataFrame(doc_analysis)
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**📊 Collection Statistics:**")
        stats_summary = {
            'Metric': [
                'Largest Document',
                'Smallest Document', 
                'Most Words',
                'Most Patterns',
                'Average Words per Document',
                'Average Patterns per Document'
            ],
            'Value': [
                doc_stats_df.loc[doc_stats_df['Content Length'].idxmax(), 'File Name'],
                doc_stats_df.loc[doc_stats_df['Content Length'].idxmin(), 'File Name'],
                doc_stats_df.loc[doc_stats_df['Word Count'].idxmax(), 'File Name'],
                doc_stats_df.loc[doc_stats_df['Pattern Count'].idxmax(), 'File Name'],
                f"{doc_stats_df['Word Count'].mean():,.0f}",
                f"{doc_stats_df['Pattern Count'].mean():.1f}"
            ]
        }
        st.dataframe(pd.DataFrame(stats_summary), use_container_width=True, hide_index=True)
    with col2:
        quality_breakdown = doc_stats_df['Content Density'].describe()
        st.markdown("**📊 Content Quality Metrics:**")
        st.write(f"• Highest content density: {quality_breakdown['max']:.1f}%")
        st.write(f"• Average content density: {quality_breakdown['mean']:.1f}%")
        st.write(f"• Lowest content density: {quality_breakdown['min']:.1f}%")
        top_docs = doc_stats_df.nlargest(3, 'Content Density')
        st.markdown("**🏆 Most Content-Rich Documents:**")
        for _, doc in top_docs.iterrows():
            st.write(f"• {doc['File Name']}")
    # Table
    st.dataframe(
        doc_stats_df.sort_values('Content Length', ascending=False),
        use_container_width=True,
        hide_index=True
    )

def show_content_analysis(engine):
    st.markdown("### 📈 Content Analysis")
    df = engine.df
    # Add more content quality metrics if needed here
    st.write("Content summary and quality metrics per document are available in the overview.")

def show_pattern_insights(engine):
    st.markdown("### 🔍 Pattern Discovery Analytics")
    if not engine.discovered_tokens:
        st.info("No patterns discovered yet."); return
    analytics_data = [
        {'Pattern': token, 'Documents': data['doc_count'], 'Total Occurrences': data['count']}
        for token, data in engine.discovered_tokens.items()
    ]
    df_token = pd.DataFrame(analytics_data).sort_values('Documents', ascending=False)
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Top Patterns:**")
        st.dataframe(df_token.head(20), use_container_width=True, hide_index=True)
    with col2:
        if len(df_token) > 0:
            fig = px.bar(
                df_token.head(10),
                x='Documents',
                y='Pattern',
                title='Top 10 Patterns by Document Count',
                orientation='h'
            )
            fig.update_layout(height=400, yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig, use_container_width=True)

def show_export_reports(engine):
    st.markdown("### 📊 Export Document & Pattern Reports")
    st.write("Use the Export tab to generate Excel reports from your search.")

def show_export_interface():
    st.subheader("📤 Export Options")
    results = st.session_state.current_results
    search_terms = st.session_state.search_terms
    search_mode = st.session_state.search_mode
    if results.empty:
        st.info("No search results to export. Perform a search first."); return
    if st.button("📊 Generate Export", type="primary"):
        export_data = create_enhanced_export(results, search_terms, search_mode)
        st.download_button(
            "💾 Download Excel Report",
            data=export_data,
            file_name=f"search_export_{search_mode.lower()}_{'_'.join(search_terms)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()