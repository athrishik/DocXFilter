#!/usr/bin/env python
# coding: utf-8

"""
DocXFilter v2.2 - Advanced Document Search & Analytics
Copyright 2025 Hrishik Kunduru. All rights reserved.

Professional multi-pattern search and analytics tool for DocXScan Excel outputs.
"""

import streamlit as st
import pandas as pd
import json
import re
import io
import html
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

# Enhanced CSS styling with professional colors
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
        padding: 2rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 20px rgba(44, 62, 80, 0.3);
    }
    
    .metric-container {
        background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
        padding: 1.2rem;
        border-radius: 8px;
        color: white;
        text-align: center;
        margin: 0.5rem 0;
        box-shadow: 0 2px 8px rgba(52, 152, 219, 0.2);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    .token-card-selected {
        background: #e8f5e8;
        border: 2px solid #27ae60;
        border-radius: 6px;
        padding: 0.75rem;
        margin: 0.3rem 0;
        box-shadow: 0 2px 6px rgba(39, 174, 96, 0.15);
    }
    
    .token-card-unselected {
        background: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 6px;
        padding: 0.75rem;
        margin: 0.3rem 0;
        transition: all 0.2s ease;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
    }
    
    .token-card-unselected:hover {
        background: #f1f3f4;
        border-color: #3498db;
        box-shadow: 0 2px 8px rgba(52, 152, 219, 0.15);
    }
    
    .search-section {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 8px;
        margin: 1rem 0;
        border-left: 4px solid #3498db;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
    }
    
    .results-header {
        background: #ecf0f1;
        padding: 1rem;
        border-radius: 6px;
        color: #2c3e50;
        font-weight: 600;
        margin: 1rem 0;
        border-left: 4px solid #e74c3c;
    }
    
    .stTabs [data-baseweb="tab-list"] { 
        gap: 2px;
        background: #ecf0f1;
        padding: 4px;
        border-radius: 6px;
        border: 1px solid #bdc3c7;
    }
    
    .stTabs [data-baseweb="tab"] { 
        height: 48px; 
        padding: 0 20px;
        border-radius: 4px;
        background: white;
        border: 1px solid transparent;
        color: #7f8c8d;
        font-weight: 500;
        transition: all 0.2s ease;
    }
    
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background: #3498db !important;
        color: white !important;
        border-color: #2980b9 !important;
        box-shadow: 0 2px 4px rgba(52, 152, 219, 0.3) !important;
        font-weight: 600 !important;
    }
    
    .stTabs [data-baseweb="tab"]:hover:not([aria-selected="true"]) {
        background: #f1f3f4;
        border-color: #bdc3c7;
        color: #2c3e50;
    }
    
    .stButton > button {
        border-radius: 4px;
        border: 1px solid #bdc3c7;
        transition: all 0.2s ease;
    }
    
    .stButton > button:hover {
        border-color: #3498db;
        color: #3498db;
    }
    
    .info-box {
        background: #e8f4fd;
        border: 1px solid #3498db;
        border-radius: 6px;
        padding: 1rem;
        margin: 1rem 0;
        color: #2c3e50;
    }
    
    .success-box {
        background: #e8f5e8;
        border: 1px solid #27ae60;
        border-radius: 6px;
        padding: 1rem;
        margin: 1rem 0;
        color: #2c3e50;
    }
    
    .warning-box {
        background: #fef9e7;
        border: 1px solid #f39c12;
        border-radius: 6px;
        padding: 1rem;
        margin: 1rem 0;
        color: #2c3e50;
    }
</style>
""", unsafe_allow_html=True)

class DocumentSearchEngine:
    """Optimized document search engine"""
    def __init__(self):
        self.df = None
        self.token_map = {}
        self.discovered_tokens = {}

    def load_data(self, df: pd.DataFrame) -> bool:
        """Load and validate Excel data"""
        try:
            required_cols = ['File Name', 'Full Contents']
            if not all(col in df.columns for col in required_cols):
                return False
            
            self.df = df.copy()
            self.df['Full Contents'] = self.df['Full Contents'].fillna('').astype(str)
            self.df['File Name'] = self.df['File Name'].astype(str)
            return True
        except Exception:
            return False

    def load_token_definitions(self, token_json: dict):
        """Load token definitions from JSON"""
        self.token_map = token_json

    @st.cache_data
    def discover_tokens(_self) -> Dict[str, Dict]:
        """Discover patterns with optimized regex"""
        if _self.df is None: 
            return {}
        
        patterns = [
            re.compile(r'<<[^>]+>>', re.IGNORECASE),
            re.compile(r'<<[^>]+\.', re.IGNORECASE),
            re.compile(r'\{[A-Z_][^}]*\}', re.IGNORECASE),
            re.compile(r'\[[A-Z_][^\]]*\]', re.IGNORECASE),
            re.compile(r'\[\[[A-Z_][^\]]*', re.IGNORECASE),
            re.compile(r'<[a-z]+>', re.IGNORECASE),
        ]
        
        all_tokens = defaultdict(lambda: {'count': 0, 'documents': set()})
        
        for _, row in _self.df.iterrows():
            content = str(row['Full Contents'])
            file_name = row['File Name']
            
            for pattern in patterns:
                for match in pattern.finditer(content):
                    token = match.group()
                    all_tokens[token]['count'] += 1
                    all_tokens[token]['documents'].add(file_name)
        
        result = {
            token: {
                'count': data['count'],
                'doc_count': len(data['documents']),
                'documents': list(data['documents'])[:20]
            } for token, data in all_tokens.items()
        }
        
        _self.discovered_tokens = result
        return result

    def search_multi(self, search_terms: List[str], mode: str) -> pd.DataFrame:
        """Optimized multi-term search"""
        if self.df is None or not search_terms:
            return pd.DataFrame()
        
        if mode == "AND":
            result_df = self.df.copy()
            for term in search_terms:
                mask = result_df['Full Contents'].str.contains(term, case=False, na=False, regex=False)
                result_df = result_df[mask]
        else:  # OR mode
            masks = [
                self.df['Full Contents'].str.contains(term, case=False, na=False, regex=False)
                for term in search_terms
            ]
            if masks:
                combined_mask = masks[0]
                for mask in masks[1:]:
                    combined_mask |= mask
                result_df = self.df[combined_mask]
            else:
                result_df = pd.DataFrame()
        
        return result_df

    def get_contexts(self, search_term: str, doc_name: str, context_length: int = 100) -> List[str]:
        """Get context around search term"""
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
            if pos == -1: 
                break
            
            context_start = max(0, pos - context_length)
            context_end = min(len(content), pos + len(search_term) + context_length)
            context = content[context_start:context_end].replace('\n', ' ').strip()
            
            if context and context not in contexts:
                contexts.append(context)
            start_pos = pos + 1
        
        return contexts

# Initialize session state efficiently
def init_session_state():
    """Initialize all session state variables"""
    defaults = {
        'search_engine': DocumentSearchEngine(),
        'data_loaded': False,
        'current_results': pd.DataFrame(),
        'search_terms': [],
        'search_mode': "AND",
        'current_search_key': "",
        'input_counter': 0,
        'tokens_loaded': False
    }
    
    for key, default_value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value

def main():
    init_session_state()
    
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>🔍 DocXFilter v2.2</h1>
        <p>Advanced multi-pattern search and analytics for DocXScan outputs</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("📁 Data Import")
        uploaded_excel = st.file_uploader("DocXScan Excel File", type=['xlsx', 'xls'])
        uploaded_tokens = st.file_uploader("Token Definitions (JSON)", type=['json'])
        
        handle_file_uploads(uploaded_excel, uploaded_tokens)
        
        if st.session_state.data_loaded:
            st.markdown("---")
            if st.button("🔄 Reset & Load New Files", type="secondary"):
                reset_application()

    # Main interface
    if st.session_state.data_loaded:
        show_main_interface()
    else:
        show_welcome_screen()

def handle_file_uploads(uploaded_excel, uploaded_tokens):
    """Handle file upload logic"""
    if uploaded_excel and not st.session_state.data_loaded:
        try:
            df = pd.read_excel(uploaded_excel, engine='openpyxl')
            if st.session_state.search_engine.load_data(df):
                st.session_state.data_loaded = True
                st.success(f"✅ Loaded {len(df)} documents")
                
                if uploaded_tokens:
                    load_tokens(uploaded_tokens)
                
                with st.spinner("🔍 Discovering patterns..."):
                    st.session_state.search_engine.discover_tokens()
                
                st.rerun()
        except Exception as e:
            st.error(f"❌ Error loading Excel file: {e}")
    
    if uploaded_tokens and st.session_state.data_loaded and not st.session_state.tokens_loaded:
        load_tokens(uploaded_tokens)

def load_tokens(uploaded_tokens):
    """Load token definitions from uploaded file"""
    try:
        uploaded_tokens.seek(0)
        token_json = json.load(uploaded_tokens)
        st.session_state.search_engine.load_token_definitions(token_json)
        st.session_state.tokens_loaded = True
        st.success(f"✅ Loaded {len(token_json)} token definitions")
        
        if token_json:
            sample_tokens = list(token_json.items())[:3]
            sample_display = ', '.join([f'`{k}`' for k, v in sample_tokens])
            st.info(f"📋 Sample: {sample_display}")
        st.rerun()
    except Exception as e:
        st.error(f"❌ Error loading tokens: {e}")

def reset_application():
    """Reset application state"""
    for key in ['data_loaded', 'current_results', 'search_terms', 'current_search_key', 'tokens_loaded']:
        st.session_state[key] = False if 'loaded' in key else (pd.DataFrame() if 'results' in key else [])
    st.session_state.search_engine = DocumentSearchEngine()
    st.rerun()

def show_welcome_screen():
    """Enhanced welcome screen"""
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div class="info-box">
        <h3>👋 Welcome to DocXFilter</h3>
        
        <strong>🎯 Advanced Document Search & Analytics Tool:</strong><br>
        🔍 Multi-pattern search with AND/OR logic<br>
        📊 Document-level analytics and insights<br>
        💡 Auto-pattern discovery<br>
        🏷️ Token management and browsing<br>
        📈 Data quality assessment<br>
        🎯 Export Excel reports<br><br>
        
        <strong>🚀 Quick Start:</strong><br>
        1. Upload DocXScan Excel file<br>
        2. Upload JSON token definitions (optional)<br>
        3. Select tokens or add custom search terms<br>
        4. Choose AND/OR search mode<br>
        5. Explore analytics and export results
        </div>
        """, unsafe_allow_html=True)

def show_main_interface():
    """Main interface with enhanced metrics"""
    engine = st.session_state.search_engine
    
    # Metrics display
    col1, col2, col3, col4 = st.columns(4)
    metrics = [
        ("📄", "Documents", len(engine.df) if engine.df is not None else 0),
        ("🔑", "Patterns Found", len(engine.discovered_tokens)),
        ("🏷️", "Imported Tokens", len(engine.token_map)),
        ("📋", "Search Results", len(st.session_state.current_results))
    ]
    
    for col, (icon, label, value) in zip([col1, col2, col3, col4], metrics):
        with col:
            st.markdown(f"""
            <div class="metric-container">
                <h3 style="margin: 0; font-size: 1.8rem;">{icon}</h3>
                <h2 style="margin: 0.5rem 0; font-size: 1.5rem;">{value:,}</h2>
                <p style="margin: 0; font-size: 0.9rem; opacity: 0.9;">{label}</p>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Tabs
    tab1, tab2, tab3 = st.tabs(["🔍 Search", "📊 Analytics", "📤 Export"])
    
    with tab1: 
        show_search_interface()
    with tab2: 
        show_analytics_interface()
    with tab3: 
        show_export_interface()

def show_search_interface():
    """Enhanced search interface"""
    st.markdown('<div class="search-section">', unsafe_allow_html=True)
    st.subheader("🔍 Multi-Pattern Search")
    
    # Search term input
    st.markdown("#### ➕ Add Search Terms")
    col1, col2, col3 = st.columns([3, 1, 1])
    
    with col1:
        new_term = st.text_input(
            "Enter text to search:", 
            placeholder="Examples: <<merge, {CLIENT_NAME}, contract, SIGNATURE...",
            key=f"search_input_{st.session_state.input_counter}"
        )
    
    with col2:
        if st.button("➕ Add Term", type="primary", use_container_width=True):
            add_search_term(new_term)
    
    with col3:
        if st.button("🧹 Clear All", use_container_width=True):
            clear_all_terms()
    
    # Token selection
    show_token_selection()
    
    # Current search terms
    if st.session_state.search_terms:
        show_current_search()
        show_search_results_summary()
    else:
        st.info("💡 Add search terms above or use 'Add Tokens from File' to start searching.")
    
    # Search results
    if not st.session_state.current_results.empty:
        show_search_results()
    
    st.markdown('</div>', unsafe_allow_html=True)

def add_search_term(term):
    """Add search term with validation"""
    if term and term.strip():
        clean_term = term.strip()
        if clean_term not in st.session_state.search_terms:
            st.session_state.search_terms.append(clean_term)
            st.session_state.input_counter += 1
            st.rerun()
        else:
            st.warning(f"'{clean_term}' is already in your search terms!")

def clear_all_terms():
    """Clear all search terms"""
    st.session_state.search_terms = []
    st.session_state.current_results = pd.DataFrame()
    st.session_state.input_counter += 1
    st.session_state.current_search_key = ""
    st.rerun()

def show_token_selection():
    """Token selection interface"""
    engine = st.session_state.search_engine
    imported_tokens = engine.token_map
    discovered_tokens = engine.discovered_tokens
    
    if not imported_tokens and not discovered_tokens:
        return
    
    with st.expander("⚡ Add Tokens from File", expanded=False):
        if imported_tokens:
            st.markdown("**📋 Available Token Definitions:**")
            st.caption(f"*{len(imported_tokens)} tokens loaded from JSON file*")
            
            # Search functionality
            search_tokens = st.text_input(
                "🔍 Search tokens:", 
                placeholder="Search by token name or description...",
                key="search_tokens_quick"
            )
            
            # Filter tokens
            filtered_tokens = filter_tokens(imported_tokens, search_tokens)
            
            if filtered_tokens:
                show_bulk_actions(filtered_tokens)
                show_token_cards(filtered_tokens)
            else:
                st.info("No tokens match your search criteria.")
            
            st.markdown("---")
        
        if discovered_tokens:
            show_discovered_patterns(discovered_tokens)

def filter_tokens(tokens, search_term):
    """Filter tokens based on search term"""
    if not search_term:
        return tokens
    
    search_lower = search_term.lower()
    return {
        token: desc for token, desc in tokens.items()
        if search_lower in token.lower() or search_lower in desc.lower()
    }

def show_bulk_actions(filtered_tokens):
    """Show bulk action buttons"""
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown(f"**{len(filtered_tokens)} tokens found:**")
    with col2:
        if st.button("➕ Add All Visible", key="add_all_filtered"):
            added_count = 0
            for token in filtered_tokens.keys():
                if token not in st.session_state.search_terms:
                    st.session_state.search_terms.append(token)
                    added_count += 1
            if added_count > 0:
                st.success(f"Added {added_count} tokens!")
                st.rerun()
    with col3:
        if st.button("❌ Clear Selected", key="clear_selected"):
            for token in list(st.session_state.search_terms):
                if token in filtered_tokens:
                    st.session_state.search_terms.remove(token)
            st.rerun()
    
    st.markdown("---")

def show_token_cards(filtered_tokens):
    """Display token cards with enhanced styling"""
    for token, description in sorted(filtered_tokens.items()):
        is_selected = token in st.session_state.search_terms
        safe_token_display = html.escape(token)
        safe_description_display = html.escape(description)
        safe_key = re.sub(r'[^a-zA-Z0-9_]', '_', token)
        
        with st.container():
            col1, col2, col3 = st.columns([4, 1, 1])
            
            with col1:
                card_class = "token-card-selected" if is_selected else "token-card-unselected"
                status_icon = "✅" if is_selected else "🏷️"
                
                st.markdown(f"""
                <div class="{card_class}">
                    <strong>{status_icon} {safe_token_display}</strong><br>
                    <small>{safe_description_display}</small>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                if is_selected:
                    if st.button("❌ Remove", key=f"remove_{safe_key}_{hash(token)}", 
                               help="Remove from search", use_container_width=True):
                        st.session_state.search_terms.remove(token)
                        st.rerun()
                else:
                    if st.button("🔍 Search", key=f"search_only_{safe_key}_{hash(token)}", 
                               help="Search this token only", use_container_width=True):
                        st.session_state.search_terms = [token]
                        st.rerun()
            
            with col3:
                if not is_selected:
                    if st.button("➕ Add", key=f"add_token_{safe_key}_{hash(token)}", 
                               help="Add to search", use_container_width=True):
                        st.session_state.search_terms.append(token)
                        st.rerun()
                else:
                    st.button("✅ Added", key=f"added_{safe_key}_{hash(token)}", 
                            disabled=True, use_container_width=True)

def show_discovered_patterns(discovered_tokens):
    """Show discovered patterns in compact format"""
    st.markdown("**🔍 Auto-Discovered Patterns:**")
    st.caption("*Top patterns found in your documents*")
    
    sorted_discovered = sorted(discovered_tokens.items(), 
                             key=lambda x: x[1]['doc_count'], reverse=True)
    cols = st.columns(4)
    
    for i, (token, data) in enumerate(sorted_discovered[:8]):
        with cols[i % 4]:
            safe_token_display = html.escape(token)[:20] + ("..." if len(token) > 20 else "")
            safe_key = re.sub(r'[^a-zA-Z0-9_]', '_', token)
            
            if st.button(f"{safe_token_display}\n({data['doc_count']} docs)", 
                       key=f"quick_discovered_{safe_key}_{i}", 
                       use_container_width=True,
                       help=f"Full token: {token}"):
                if token not in st.session_state.search_terms:
                    st.session_state.search_terms.append(token)
                    st.rerun()

def show_current_search():
    """Display current search terms and controls"""
    st.markdown("#### 🔍 Current Search Terms")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        mode = st.radio(
            "Search Mode:", ["AND", "OR"], 
            index=0 if st.session_state.search_mode == "AND" else 1,
            horizontal=True,
            help="AND: All terms must be present | OR: Any term present"
        )
        st.session_state.search_mode = mode
    
    with col2:
        if st.button("🔍 Search Now", type="primary", use_container_width=True):
            perform_multi_search()
    
    # Display active terms without HTML
    st.markdown("**Active Search Terms:**")
    for i, term in enumerate(st.session_state.search_terms):
        col1, col2 = st.columns([6, 1])
        with col1:
            # Use st.code to safely display any characters
            st.code(term, language=None)
        with col2:
            if st.button("❌", key=f"remove_term_{i}", help=f"Remove '{term}'"):
                st.session_state.search_terms.pop(i)
                st.session_state.current_search_key = ""
                st.rerun()
    
    # Auto-search on changes
    key = f"{mode}:{'|'.join(sorted(st.session_state.search_terms))}"
    if key != st.session_state.current_search_key:
        perform_multi_search()

def show_search_results_summary():
    """Show search results summary without HTML rendering"""
    if not st.session_state.current_results.empty:
        results_count = len(st.session_state.current_results)
        search_mode = st.session_state.search_mode
        search_terms = st.session_state.search_terms
        
        st.success(f"✅ Found {results_count} documents ({search_mode} search)")
        with st.expander("Search Details", expanded=False):
            st.write("**Search Terms:**")
            for term in search_terms:
                st.code(term, language=None)
            
    elif st.session_state.search_terms and st.session_state.current_search_key:
        search_mode = st.session_state.search_mode
        search_terms = st.session_state.search_terms
        
        st.warning(f"⚠️ No documents found ({search_mode} search)")
        with st.expander("Search Details", expanded=False):
            st.write("**Search Terms:**")
            for term in search_terms:
                st.code(term, language=None)

def perform_multi_search():
    """Execute multi-term search"""
    terms = st.session_state.search_terms
    mode = st.session_state.search_mode
    engine = st.session_state.search_engine
    
    result_df = engine.search_multi(terms, mode)
    st.session_state.current_results = result_df
    st.session_state.current_search_key = f"{mode}:{'|'.join(sorted(terms))}"

def show_search_results():
    """Enhanced search results display without HTML issues"""
    results = st.session_state.current_results
    search_terms = st.session_state.search_terms
    search_mode = st.session_state.search_mode
    
    st.markdown("---")
    st.subheader(f"📋 Results ({len(results)} documents)")
    st.write(f"**Search Mode:** {search_mode}")
    
    # Export button
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("📤 Export Results", use_container_width=True):
            export_data = create_enhanced_export(results, search_terms, search_mode)
            filename = f"docxfilter_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.download_button(
                "💾 Download Excel", data=export_data, file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # Results table
    display_results_table(results, search_terms)
    
    # Document preview
    if len(results) > 0:
        show_document_preview(results, search_terms)

def display_results_table(results, search_terms):
    """Display results table with match summary"""
    display_cols = ['File Name']
    for col in ['Size (KB)', 'Content Length (chars)']:
        if col in results.columns:
            display_cols.append(col)
    
    if not results.empty:
        results_with_matches = results.copy()
        match_summaries = []
        
        for _, row in results_with_matches.iterrows():
            content = str(row['Full Contents']).lower()
            matches = [f"{term}:{content.count(term.lower())}" 
                      for term in search_terms if content.count(term.lower()) > 0]
            match_summaries.append(" | ".join(matches))
        
        results_with_matches['Match Summary'] = match_summaries
        display_cols.append('Match Summary')
        
        st.dataframe(
            results_with_matches[display_cols],
            use_container_width=True,
            hide_index=True,
            height=300
        )

def show_document_preview(results, search_terms):
    """Enhanced document preview without HTML issues"""
    st.subheader("📖 Document Preview")
    selected_doc = st.selectbox(
        "Select document to preview:", results['File Name'].tolist(),
        help="Choose a document to see all matched terms in context")
    
    if selected_doc:
        engine = st.session_state.search_engine
        doc_row = engine.df[engine.df['File Name'] == selected_doc]
        
        if not doc_row.empty:
            content = str(doc_row['Full Contents'].iloc[0])
            
            # Document info
            col1, col2 = st.columns(2)
            with col1: 
                st.write(f"**📄 {selected_doc}**")
            with col2:
                st.write("**🔍 Term Occurrences:**")
                for term in search_terms:
                    count = content.lower().count(term.lower())
                    st.write(f"• {count} times")
                    st.code(term, language=None)
            
            # Context preview
            show_context_preview(engine, selected_doc, search_terms)
            
            # Full document content
            show_document_content(content, search_terms)

def show_context_preview(engine, doc_name, search_terms):
    """Show context preview without HTML rendering issues"""
    st.write("**📋 Context Preview for Each Term:**")
    
    for i, term in enumerate(search_terms):
        with st.expander(f"🔍 Contexts for term {i+1}", expanded=(i == 0)):
            st.code(term, language=None)
            
            contexts = engine.get_contexts(term, doc_name)
            
            if contexts:
                for j, context in enumerate(contexts, 1):
                    st.write(f"**Context {j}:**")
                    # Highlight the term in the context using text highlighting
                    lines = context.split('\n')
                    for line in lines:
                        if term.lower() in line.lower():
                            st.info(line.strip())
                        else:
                            st.write(line.strip())
                    st.markdown("---")
            else:
                st.write("No contexts found for this term in this document.")

def show_document_content(content, search_terms):
    """Show full document content safely"""
    st.markdown("#### 📝 Full Document Content")
    
    # Use text area for safe display of any content
    st.text_area(
        "Document content (use Ctrl+F to search for terms)",
        value=content,
        height=400,
        help="Use your browser's search (Ctrl+F) to find specific terms in this content"
    )
    
    # Show search term summary
    if search_terms:
        st.write("**Search Terms in this document:**")
        for i, term in enumerate(search_terms, 1):
            count = content.lower().count(term.lower())
            col1, col2 = st.columns([3, 1])
            with col1:
                st.code(term, language=None)
            with col2:
                st.metric("Occurrences", count)

def create_enhanced_export(results_df: pd.DataFrame, search_terms: List[str], search_mode: str) -> bytes:
    """Create enhanced Excel export with error handling"""
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if results_df.empty:
                # Create summary sheet for no results
                summary_data = {
                    'Search Mode': [search_mode],
                    'Search Terms': [' | '.join(search_terms) if search_terms else 'None'],
                    'Results Found': [0],
                    'Export Date': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                    'Generated By': ["DocXFilter v2.2"],
                    'Note': ['No documents found matching search criteria']
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Search Summary', index=False)
                
                # Create empty results sheet
                empty_results = pd.DataFrame({'File Name': [], 'Note': []})
                empty_results.to_excel(writer, sheet_name='Search Results', index=False)
            else:
                # Create results sheet with data
                export_df = results_df.copy()
                
                # Add match columns safely
                try:
                    for term in search_terms:
                        if 'Full Contents' in export_df.columns:
                            export_df[f"'{term}' Count"] = export_df['Full Contents'].str.lower().str.count(term.lower())
                        else:
                            export_df[f"'{term}' Count"] = 0
                    
                    if search_terms and 'Full Contents' in export_df.columns:
                        export_df['Total Matches'] = sum([export_df[f"'{term}' Count"] for term in search_terms])
                    else:
                        export_df['Total Matches'] = 0
                        
                except Exception:
                    pass
                
                # Determine export columns
                basic_cols = ['File Name']
                optional_cols = ['Size (KB)', 'Content Length (chars)']
                match_cols = [f"'{term}' Count" for term in search_terms if f"'{term}' Count" in export_df.columns]
                
                export_cols = basic_cols + [col for col in optional_cols if col in export_df.columns] + match_cols
                if 'Total Matches' in export_df.columns:
                    export_cols.append('Total Matches')
                
                # Export main results
                export_df[export_cols].to_excel(writer, sheet_name='Search Results', index=False)
                
                # Create summary sheet
                summary_data = {
                    'Search Mode': [search_mode],
                    'Search Terms': [' | '.join(search_terms) if search_terms else 'None'],
                    'Number of Terms': [len(search_terms)],
                    'Results Found': [len(results_df)],
                    'Export Date': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                    'Generated By': ['DocXFilter v2.2']
                }
                
                # Add individual terms
                for i, term in enumerate(search_terms, 1):
                    summary_data[f'Term {i}'] = [term]
                
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Search Summary', index=False)
        
        output.seek(0)
        return output.read()
        
    except Exception as e:
        st.error(f"Export error: {e}")
        
        # Fallback: create simple error report
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            error_data = {
                'Export Status': ['Error'],
                'Error Message': [str(e)],
                'Search Terms': [' | '.join(search_terms) if search_terms else 'None'],
                'Results Count': [len(results_df) if not results_df.empty else 0],
                'Export Date': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
            }
            error_df = pd.DataFrame(error_data)
            error_df.to_excel(writer, sheet_name='Export Error', index=False)
        
        output.seek(0)
        return output.read()

def show_analytics_interface():
    """Enhanced analytics interface"""
    st.subheader("📊 Document Analytics Dashboard")
    engine = st.session_state.search_engine
    
    if engine.df is None or len(engine.df) == 0:
        st.info("📄 No documents available for analytics.")
        return
    
    tab1, tab2 = st.tabs(["📄 Document Overview", "🔍 Pattern Insights"])
    
    with tab1:
        show_document_overview(engine)
    with tab2:
        show_pattern_insights(engine)

def show_document_overview(engine):
    """Enhanced document overview"""
    st.markdown("### 📄 Document Collection Overview")
    df = engine.df
    
    # Key metrics
    col1, col2, col3, col4 = st.columns(4)
    content_lengths = df['Full Contents'].str.len()
    
    metrics = [
        ("📄", "Total Documents", len(df)),
        ("💾", "Total Size (KB)", f"{df.get('Size (KB)', pd.Series()).sum():,.1f}" if 'Size (KB)' in df.columns else "N/A"),
        ("📝", "Total Characters", f"{content_lengths.sum():,}"),
        ("📊", "Avg Doc Length", f"{content_lengths.mean():,.0f} chars")
    ]
    
    for col, (icon, label, value) in zip([col1, col2, col3, col4], metrics):
        with col:
            st.markdown(f"""
            <div style="
                background: #3498db;
                padding: 1rem;
                border-radius: 6px;
                color: white;
                text-align: center;
                margin: 0.5rem 0;
                box-shadow: 0 2px 4px rgba(52, 152, 219, 0.2);
            ">
                <h3 style="margin: 0; font-size: 1.5rem;">{icon}</h3>
                <h4 style="margin: 0.5rem 0; font-size: 1.2rem;">{value}</h4>
                <p style="margin: 0; font-size: 0.85rem; opacity: 0.9;">{label}</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Distribution charts
    col1, col2 = st.columns(2)
    
    with col1:
        if 'Size (KB)' in df.columns and df['Size (KB)'].notna().any():
            fig = px.histogram(
                x=df['Size (KB)'].dropna(),
                nbins=25, 
                title='📦 Document Size Distribution',
                labels={'x': 'File Size (KB)', 'y': 'Number of Documents'},
                color_discrete_sequence=['#3498db']
            )
            fig.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("📊 No file size information available")
    
    with col2:
        fig = px.histogram(
            x=content_lengths,
            nbins=25,
            title='📝 Content Length Distribution',
            labels={'x': 'Content Length (characters)', 'y': 'Number of Documents'},
            color_discrete_sequence=['#2980b9']
        )
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    # Document statistics
    show_document_statistics(df, content_lengths)

def show_document_statistics(df, content_lengths):
    """Show detailed document statistics"""
    st.markdown("### 📋 Document Statistics Summary")
    
    # Calculate statistics
    doc_analysis = []
    for _, row in df.iterrows():
        content = str(row['Full Contents'])
        word_count = len(content.split()) if content.strip() else 0
        line_count = content.count('\n') + 1 if content.strip() else 0
        
        # Pattern counts
        pattern_counts = [
            len(re.findall(r'<<[^>]*>>', content)),
            len(re.findall(r'\{[^}]*\}', content)),
            len(re.findall(r'\[[^\]]*\]', content))
        ]
        total_patterns = sum(pattern_counts)
        
        doc_analysis.append({
            'File Name': row['File Name'],
            'Size (KB)': row.get('Size (KB)', 0),
            'Content Length': len(content),
            'Word Count': word_count,
            'Line Count': line_count,
            'Pattern Count': total_patterns,
            'Content Density': word_count / max(len(content), 1) * 100
        })
    
    doc_stats_df = pd.DataFrame(doc_analysis)
    
    # Summary statistics
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**📊 Collection Statistics:**")
        stats_summary = {
            'Metric': [
                'Largest Document', 'Smallest Document', 'Most Words',
                'Most Patterns', 'Avg Words/Doc', 'Avg Patterns/Doc'
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
    
    # Full statistics table
    st.dataframe(
        doc_stats_df.sort_values('Content Length', ascending=False),
        use_container_width=True,
        hide_index=True
    )

def show_pattern_insights(engine):
    """Enhanced pattern insights"""
    st.markdown("### 🔍 Pattern Discovery Analytics")
    
    imported_tokens = engine.token_map
    discovered_tokens = engine.discovered_tokens
    
    if not imported_tokens and not discovered_tokens:
        st.info("🔍 No patterns discovered yet.")
        return
    
    col1, col2 = st.columns(2)
    
    # Imported tokens
    with col1:
        st.markdown("#### 🏷️ Imported Tokens")
        if imported_tokens:
            imported_data = [
                {
                    'Token': token, 
                    'Description': desc,
                    'Available': '✅'
                }
                for token, desc in imported_tokens.items()
            ]
            df_imported = pd.DataFrame(imported_data)
            st.dataframe(df_imported, use_container_width=True, hide_index=True)
        else:
            st.info("No imported tokens loaded.")
    
    # Discovered patterns
    with col2:
        st.markdown("#### 🔍 Top Discovered Patterns")
        if discovered_tokens:
            discovered_data = [
                {
                    'Pattern': token,
                    'Documents': data['doc_count'],
                    'Occurrences': data['count']
                }
                for token, data in discovered_tokens.items()
            ]
            df_discovered = pd.DataFrame(discovered_data).sort_values('Documents', ascending=False)
            st.dataframe(df_discovered.head(15), use_container_width=True, hide_index=True)
        else:
            st.info("No patterns discovered.")
    
    # Visualization
    if discovered_tokens:
        st.markdown("#### 📊 Pattern Distribution")
        analytics_data = [
            {
                'Pattern': token,
                'Documents': data['doc_count'],
                'Total Occurrences': data['count'],
                'Type': 'Discovered'
            }
            for token, data in discovered_tokens.items()
        ]
        
        df_all = pd.DataFrame(analytics_data).sort_values('Documents', ascending=False)
        
        if len(df_all) > 0:
            fig = px.scatter(
                df_all.head(20),
                x='Documents',
                y='Total Occurrences',
                color='Type',
                hover_data=['Pattern'],
                title='Top 20 Patterns: Document Count vs Total Occurrences',
                labels={'Documents': 'Number of Documents', 'Total Occurrences': 'Total Occurrences'},
                color_discrete_sequence=['#3498db', '#2980b9']
            )
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)

def show_export_interface():
    """Enhanced export interface"""
    st.subheader("📤 Export Options")
    results = st.session_state.current_results
    search_terms = st.session_state.search_terms
    search_mode = st.session_state.search_mode
    
    if results.empty:
        st.markdown("""
        <div class="warning-box">
            <h4>📄 No search results to export</h4>
            <p>Perform a search first to generate exportable results.</p>
        </div>
        """, unsafe_allow_html=True)
        return
    
    # Export summary
    st.markdown(f"""
    <div class="success-box">
        <h4>📊 Export Summary</h4>
        <p><strong>Results:</strong> {len(results)} documents found</p>
        <p><strong>Search Mode:</strong> {search_mode}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Show search terms safely
    st.write("**Search Terms:**")
    for term in search_terms:
        st.code(term, language=None)
    
    # Export button with error handling
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("📊 Generate Excel Export", type="primary", use_container_width=True):
            try:
                with st.spinner("Generating Excel export..."):
                    export_data = create_enhanced_export(results, search_terms, search_mode)
                    
                if export_data:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"docxfilter_export_{timestamp}.xlsx"
                    
                    st.download_button(
                        "💾 Download Excel Report",
                        data=export_data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    st.success("✅ Export ready for download!")
                else:
                    st.error("❌ Failed to generate export. Please try again.")
                    
            except Exception as e:
                st.error(f"❌ Export failed: {str(e)}")
                st.info("💡 Try reducing the number of search results or check your data.")

if __name__ == "__main__":
    main()
