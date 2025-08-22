#!/usr/bin/env python
# coding: utf-8

"""
DocXFilter v3.0 - Optimized Advanced Document Search & Analytics
Copyright 2025 Hrishik Kunduru. All rights reserved.

Professional multi-pattern search and analytics tool for DocXScan Excel outputs.
OPTIMIZATIONS: Improved performance, reduced redundancy, better caching, modern UI colors
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
from typing import Dict, List, Optional, Tuple
from functools import lru_cache

# Streamlit page
st.set_page_config(
    page_title="DocXFilter v3.0",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e40af 0%, #3730a3 100%);
        padding: 2rem;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(30, 64, 175, 0.25);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    .search-section {
        background: linear-gradient(135deg, #f8fafc 5%, #e2e8f0 100%);
        padding: 2rem;
        border-radius: 16px;
        margin: 1rem 0;
        border: 1px solid #cbd5e1;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.06);
        color: #334155;
    }
    
    .search-clear-btn button {
        background: linear-gradient(135deg, #dc2626 0%, #b91c1c 100%) !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
    }
    
    .search-clear-btn button:hover {
        background: linear-gradient(135deg, #b91c1c 0%, #991b1b 100%) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(220, 38, 38, 0.3) !important;
    }
    
    .stTabs [data-baseweb="tab-list"] { 
        gap: 4px;
        background: linear-gradient(135deg, #f1f5f9 0%, #e2e8f0 100%);
        padding: 8px;
        border-radius: 12px;
        border: 1px solid #cbd5e1;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
    }
    
    .stTabs [data-baseweb="tab"] { 
        height: 52px; 
        padding: 0 24px;
        border-radius: 10px;
        background: white;
        border: 1px solid #e2e8f0;
        color: #64748b;
        font-weight: 500;
        transition: all 0.3s ease;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.04);
    }
    
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #1e40af 0%, #3730a3 100%) !important;
        color: white !important;
        border-color: #1d4ed8 !important;
        box-shadow: 0 4px 20px rgba(30, 64, 175, 0.25) !important;
        font-weight: 600 !important;
        transform: translateY(-2px) !important;
    }
</style>
""", unsafe_allow_html=True)

class DocumentSearchEngine:
    """Streamlined document search engine"""
    
    def __init__(self):
        self.df: Optional[pd.DataFrame] = None
        self.token_map: Dict[str, str] = {}
        self.discovered_tokens: Dict[str, Dict] = {}
        self._content_cache: Dict[str, str] = {}
        self._compiled_patterns = self._compile_patterns()

    @staticmethod
    @lru_cache(maxsize=1)
    def _compile_patterns() -> List[re.Pattern]:
        """Pre-compile regex patterns for better performance"""
        return [
            re.compile(r'<<[^>]+>>', re.IGNORECASE),
            re.compile(r'<<[^>]+\.', re.IGNORECASE),
            re.compile(r'\{[A-Z_][^}]*\}', re.IGNORECASE),
            re.compile(r'\[[A-Z_][^\]]*\]', re.IGNORECASE),
            re.compile(r'\[\[[A-Z_][^\]]*', re.IGNORECASE),
            re.compile(r'<[a-z]+>', re.IGNORECASE),
        ]

    def load_data(self, df: pd.DataFrame) -> bool:
        """Load and validate Excel data"""
        try:
            required_cols = ['File Name', 'Full Contents']
            if not all(col in df.columns for col in required_cols):
                return False
            
            self.df = df.copy()
            self.df['Full Contents'] = self.df['Full Contents'].fillna('').astype('string')
            self.df['File Name'] = self.df['File Name'].astype('string')
            self._content_cache = dict(zip(self.df['File Name'], self.df['Full Contents']))
            return True
        except Exception:
            return False

    def load_token_definitions(self, token_json: Dict[str, str]) -> None:
        """Load token definitions from JSON"""
        self.token_map = token_json

    @st.cache_data(ttl=3600)
    def discover_tokens(_self) -> Dict[str, Dict]:
        """Optimized token discovery"""
        if _self.df is None: 
            return {}
        
        all_tokens = defaultdict(lambda: {'count': 0, 'documents': set()})
        
        for file_name, content in _self._content_cache.items():
            for pattern in _self._compiled_patterns:
                matches = pattern.findall(str(content))
                for match in matches:
                    all_tokens[match]['count'] += 1
                    all_tokens[match]['documents'].add(file_name)
        
        result = {
            token: {
                'count': data['count'],
                'doc_count': len(data['documents']),
                'documents': list(data['documents'])[:20]
            } for token, data in all_tokens.items()
        }
        
        _self.discovered_tokens = result
        return result

    @lru_cache(maxsize=100)
    def _search_single_term(self, term: str, mode: str) -> Tuple[str, ...]:
        """Cached single term search"""
        if self.df is None:
            return ()
        
        mask = self.df['Full Contents'].str.contains(term, case=False, na=False, regex=False)
        return tuple(self.df[mask]['File Name'].tolist())

    def search_multi(self, search_terms: List[str], mode: str) -> pd.DataFrame:
        """Optimized multi-term search"""
        if self.df is None or not search_terms:
            return pd.DataFrame()
        
        term_results = [set(self._search_single_term(term, mode)) for term in search_terms]
        
        if mode == "AND":
            result_files = set.intersection(*term_results) if term_results else set()
        else:  # OR mode
            result_files = set.union(*term_results) if term_results else set()
        
        return self.df[self.df['File Name'].isin(result_files)].copy() if result_files else pd.DataFrame()

    @lru_cache(maxsize=50)
    def get_contexts(self, search_term: str, doc_name: str, context_length: int = 100) -> Tuple[str, ...]:
        """Optimized context extraction"""
        content = self._content_cache.get(doc_name, '')
        if not content or not search_term.strip():
            return ()
        
        contexts = []
        search_lower = search_term.strip().lower()
        content_lower = content.lower()
        start_pos = 0
        
        while len(contexts) < 5:
            pos = content_lower.find(search_lower, start_pos)
            if pos == -1: 
                break
            
            context_start = max(0, pos - context_length)
            context_end = min(len(content), pos + len(search_term) + context_length)
            
            # Try to break at word boundaries
            if context_start > 0:
                for i in range(context_start, min(context_start + 50, len(content))):
                    if content[i] in ' \n\t.!?;':
                        context_start = i + 1
                        break
            
            context = content[context_start:context_end].replace('\n', ' ').strip()
            if context and len(context) > 10 and context not in contexts:
                prefix = "..." if context_start > 0 else ""
                suffix = "..." if context_end < len(content) else ""
                contexts.append(f"{prefix}{context}{suffix}")
                
            start_pos = pos + 1
        
        return tuple(contexts)

    def get_all_matched_contexts(self, search_terms: List[str], doc_name: str, context_length: int = 150) -> str:
        """Get all matched contexts for Excel export"""
        content = self._content_cache.get(doc_name, '')
        if not content:
            return ""
        
        all_contexts = []
        for term in search_terms:
            contexts = self.get_contexts(term, doc_name, context_length)
            for context in contexts:
                all_contexts.append(f"[{term}]: {context}")
        
        return " | ".join(all_contexts)

# Optimized session state initialization
@st.cache_resource
def get_search_engine():
    """Cached search engine instance"""
    return DocumentSearchEngine()

def init_session_state():
    """Streamlined session state initialization"""
    defaults = {
        'search_engine': get_search_engine(),
        'data_loaded': False,
        'current_results': pd.DataFrame(),
        'search_terms': [],
        'search_mode': "AND",
        'current_search_key': "",
        'input_counter': 0,
        'tokens_loaded': False,
        'last_file_hash': None  # Track file changes
    }
    
    for key, default_value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default_value

def main():
    init_session_state()
    
    # Enhanced header with version info
    st.markdown("""
    <div class="main-header">
        <h1>🔍 DocXFilter v3.0</h1>
        <p>Optimized multi-pattern search and analytics • Enhanced performance & modern UI</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar with improved organization
    with st.sidebar:
        st.header("📁 Data Import")
        
        # File uploaders with better UX
        uploaded_excel = st.file_uploader(
            "DocXScan Excel File", 
            type=['xlsx', 'xls'],
            help="Upload your DocXScan output file"
        )
        
        uploaded_tokens = st.file_uploader(
            "Token Definitions (JSON)", 
            type=['json'],
            help="Optional: Upload token definitions for enhanced search"
        )
        
        handle_file_uploads(uploaded_excel, uploaded_tokens)
        
        if st.session_state.data_loaded:
            st.markdown("---")
            st.success("✅ Data loaded successfully!")
            
            # Show data summary
            engine = st.session_state.search_engine
            if engine.df is not None:
                st.metric("Documents", len(engine.df))
                if st.session_state.tokens_loaded:
                    st.metric("Token Definitions", len(engine.token_map))
            
            if st.button("🔄 Reset & Load New Files", type="secondary"):
                reset_application()

    # Main interface
    if st.session_state.data_loaded:
        show_main_interface()
    else:
        show_welcome_screen()

def handle_file_uploads(uploaded_excel, uploaded_tokens):
    """Optimized file upload handling"""
    if uploaded_excel and not st.session_state.data_loaded:
        # Create file hash to detect changes
        file_hash = hash(uploaded_excel.getvalue())
        
        if st.session_state.last_file_hash != file_hash:
            try:
                with st.spinner("🔄 Loading Excel file..."):
                    df = pd.read_excel(uploaded_excel, engine='openpyxl')
                    
                if st.session_state.search_engine.load_data(df):
                    st.session_state.data_loaded = True
                    st.session_state.last_file_hash = file_hash
                    
                    # Clear search cache when new data is loaded
                    st.session_state.search_engine._search_single_term.cache_clear()
                    st.session_state.search_engine.get_contexts.cache_clear()
                    
                    if uploaded_tokens:
                        load_tokens(uploaded_tokens)
                    
                    # Discover tokens asynchronously
                    with st.spinner("🔍 Discovering patterns..."):
                        st.session_state.search_engine.discover_tokens()
                    
                    st.rerun()
                else:
                    st.error("❌ Invalid Excel format. Please ensure file has 'File Name' and 'Full Contents' columns.")
                    
            except Exception as e:
                st.error(f"❌ Error loading Excel file: {e}")
                st.info("💡 Ensure the file is a valid DocXScan output with proper formatting.")
    
    if uploaded_tokens and st.session_state.data_loaded and not st.session_state.tokens_loaded:
        load_tokens(uploaded_tokens)

def load_tokens(uploaded_tokens):
    """Optimized token loading"""
    try:
        uploaded_tokens.seek(0)
        token_json = json.load(uploaded_tokens)
        st.session_state.search_engine.load_token_definitions(token_json)
        st.session_state.tokens_loaded = True
        
        # Show preview without blocking UI
        if token_json:
            sample_count = min(3, len(token_json))
            sample_tokens = list(token_json.keys())[:sample_count]
            sample_display = ', '.join([f'`{k}`' for k in sample_tokens])
            remaining = len(token_json) - sample_count
            
            if remaining > 0:
                st.success(f"✅ Loaded {len(token_json)} tokens: {sample_display} (+{remaining} more)")
            else:
                st.success(f"✅ Loaded {len(token_json)} tokens: {sample_display}")
                
        st.rerun()
    except Exception as e:
        st.error(f"❌ Error loading tokens: {e}")

def reset_application():
    """Optimized application reset"""
    # Clear all caches
    if hasattr(st.session_state.search_engine, '_search_single_term'):
        st.session_state.search_engine._search_single_term.cache_clear()
    if hasattr(st.session_state.search_engine, 'get_contexts'):
        st.session_state.search_engine.get_contexts.cache_clear()
    
    # Clear the cached search engine function
    get_search_engine.clear()
    
    # Reset state
    st.session_state.update({
        'data_loaded': False,
        'current_results': pd.DataFrame(),
        'search_terms': [],
        'current_search_key': "",
        'tokens_loaded': False,
        'last_file_hash': None,
        'search_engine': get_search_engine()
    })
    st.rerun()

def show_welcome_screen():
    """Compact welcome screen"""
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div class="info-box">
        <h3>👋 Welcome to DocXFilter v3.0</h3>
        
        <strong>🎯 Core Features:</strong><br>
        🔍 Multi-pattern search (AND/OR logic)<br>
        📊 Document analytics with line numbers<br>
        💡 Auto-pattern discovery<br>
        🏷️ Token management system<br>
        📤 Enhanced Excel exports<br><br>
        
        <strong>🚀 Quick Start:</strong><br>
        1. Upload DocXScan Excel file<br>
        2. Optional: Upload JSON token definitions<br>
        3. Type search terms and press Enter<br>
        4. View results with line numbers<br>
        5. Export findings to Excel
        </div>
        """, unsafe_allow_html=True)

def show_main_interface():
    """Optimized main interface with better performance"""
    engine = st.session_state.search_engine
    
    # Optimized metrics display
    metrics_data = get_metrics_data(engine)
    display_metrics(metrics_data)
    
    st.markdown("---")
    
    # Optimized tabs with lazy loading
    tab1, tab2, tab3 = st.tabs(["🔍 Search", "📊 Analytics", "📤 Export"])
    
    with tab1: 
        show_search_interface()
    with tab2: 
        show_analytics_interface()
    with tab3: 
        show_export_interface()

def get_metrics_data(engine) -> Dict[str, int]:
    """Cached metrics calculation"""
    return {
        'documents': len(engine.df) if engine.df is not None else 0,
        'patterns': len(engine.discovered_tokens),
        'tokens': len(engine.token_map),
        'results': len(st.session_state.current_results)
    }

def display_metrics(metrics: Dict[str, int]):
    """Optimized metrics display with distinct colors"""
    col1, col2, col3, col4 = st.columns(4)
    
    # Define sleek color schemes for each metric card
    metric_configs = [
        ("📄", "Documents", metrics['documents'], "linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%)", "rgba(30, 58, 138, 0.25)"),
        ("🔑", "Patterns Found", metrics['patterns'], "linear-gradient(135deg, #059669 0%, #10b981 100%)", "rgba(5, 150, 105, 0.25)"),
        ("🏷️", "Imported Tokens", metrics['tokens'], "linear-gradient(135deg, #dc2626 0%, #ef4444 100%)", "rgba(220, 38, 38, 0.25)"),
        ("📋", "Search Results", metrics['results'], "linear-gradient(135deg, #7c3aed 0%, #8b5cf6 100%)", "rgba(124, 58, 237, 0.25)")
    ]
    
    for col, (icon, label, value, bg_gradient, shadow_color) in zip([col1, col2, col3, col4], metric_configs):
        with col:
            st.markdown(f"""
            <div style="
                background: {bg_gradient};
                padding: 1.5rem;
                border-radius: 12px;
                color: white;
                text-align: center;
                margin: 0.5rem 0;
                box-shadow: 0 4px 20px {shadow_color};
                border: 1px solid rgba(255, 255, 255, 0.15);
                transition: transform 0.2s ease;
            " onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 6px 25px {shadow_color.replace('0.25', '0.35')}'" 
               onmouseout="this.style.transform='translateY(0px)'; this.style.boxShadow='0 4px 20px {shadow_color}'">
                <h3 style="margin: 0; font-size: 1.8rem;">{icon}</h3>
                <h2 style="margin: 0.5rem 0; font-size: 1.5rem;">{value:,}</h2>
                <p style="margin: 0; font-size: 0.9rem; opacity: 0.9;">{label}</p>
            </div>
            """, unsafe_allow_html=True)

def show_search_interface():
    """Streamlined search interface"""
    st.markdown('<div class="search-section">', unsafe_allow_html=True)
    st.markdown("### 🔍 Multi-Pattern Search")
    
    # Single form-based input with Enter key support
    col1, col2 = st.columns([5, 1])
    with col1:
        with st.form(key="search_term_form", clear_on_submit=True):
            new_term = st.text_input(
                "Enter search term:", 
                placeholder="e.g., <<merge>>, {CLIENT_NAME}, contract, SIGNATURE... (Press Enter to add)",
                key=f"search_input_form_{st.session_state.input_counter}",
                help="Type search terms and press Enter to add them"
            )
            # This creates the functional "Add Term" button
            submitted = st.form_submit_button("Add Term", type="primary")
            if submitted and new_term:
                add_search_term(new_term)
    
    with col2:
        st.markdown('<div class="search-clear-btn">', unsafe_allow_html=True)
        if st.button("🧹 Clear All", key="clear_search_terms_button", use_container_width=True):
            clear_all_terms()
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Token selection
    show_token_selection()
    
    # Current search and results
    if st.session_state.search_terms:
        show_current_search()
        show_search_results_summary()
        if not st.session_state.current_results.empty:
            show_search_results()
    else:
        st.info("💡 Add search terms above or select from available tokens to start searching.")
    
    st.markdown('</div>', unsafe_allow_html=True)

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
    """Streamlined document preview"""
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
                    st.write(f"• **{term}**: {count} times")
            
            # Main document content with highlighting
            show_document_content_highlighted(content, search_terms)
            
            # Optional context preview
            st.markdown("---")
            if st.checkbox("📋 Show Individual Term Contexts", key="show_contexts"):
                show_context_preview_highlighted(engine, selected_doc, search_terms)

def show_context_preview_highlighted(engine, doc_name, search_terms):
    """Show context preview with completely safe text-only highlighting"""
    st.write("**📋 Context Preview for Each Term:**")
    
    # Add debug toggle
    debug_mode = st.checkbox("🔧 Show Debug Info", key="debug_highlighting")
    
    if debug_mode:
        content = engine._content_cache.get(doc_name, '') if hasattr(engine, '_content_cache') else ""
        st.write("### 🔧 Debug Information")
        
        for term in search_terms:
            st.write(f"**Checking term: `{term}`**")
            
            # Check exact matches
            exact_count = content.lower().count(term.lower())
            st.write(f"- Exact matches found: {exact_count}")
            
            # Show first few characters around matches
            if exact_count > 0:
                pos = content.lower().find(term.lower())
                if pos >= 0:
                    start = max(0, pos - 30)
                    end = min(len(content), pos + len(term) + 30)
                    context_preview = content[start:end]
                    st.code(f"Context: ...{context_preview}...")
            st.write("---")
    
    # Use tabs instead of nested expanders - NO HTML highlighting
    if len(search_terms) > 1:
        tabs = st.tabs([f"🔍 {term[:20]}{'...' if len(term) > 20 else ''}" for term in search_terms])
        
        for i, (term, tab) in enumerate(zip(search_terms, tabs)):
            with tab:
                show_contexts_for_term(engine, doc_name, term)
    else:
        # If only one term, show directly without tabs
        if search_terms:
            st.write(f"**🔍 Contexts for: {search_terms[0]}**")
            show_contexts_for_term(engine, doc_name, search_terms[0])

def show_contexts_for_term(engine, doc_name, term):
    """Show contexts for a single term using safe highlighting like v2.2"""
    contexts = engine.get_contexts(term, doc_name)
    
    if contexts:
        for j, context in enumerate(contexts, 1):
            st.write(f"**Context {j}:**")
            
            # Use the same safe approach as your v2.2 code
            # Highlight by showing lines differently instead of HTML
            lines = context.split('\n') if '\n' in context else [context]
            
            for line in lines:
                line = line.strip()
                if line and term.lower() in line.lower():
                    # Use st.info for lines containing the term (like your v2.2 code)
                    st.info(line)
                elif line:
                    # Regular text for other lines
                    st.write(line)
            
            if j < len(contexts):
                st.markdown("---")
    else:
        st.write("No contexts found for this term in this document.")
        
        # Show why no contexts were found
        content = engine._content_cache.get(doc_name, '') if hasattr(engine, '_content_cache') else ""
        if content:
            exact_matches = content.lower().count(term.lower())
            if exact_matches == 0:
                st.write(f"🔍 Term '{term}' not found in document")
            else:
                st.write(f"✅ Term found {exact_matches} times, but context extraction failed")

def show_document_content_highlighted(content, search_terms):
    """Streamlined document content with ultra-compact line preview"""
    st.markdown("#### 📝 Full Document Content")
    
    # Ultra-compact term location preview - single line
    if search_terms and content:
        lines = content.split('\n')
        line_details = []
        
        # Find matching lines efficiently
        for line_num, line in enumerate(lines):
            for term in search_terms:
                if term and term.strip() and term.strip().lower() in line.lower():
                    line_details.append(line_num + 1)
                    break  # Only need to know the line matches, not which terms
        
        if line_details:
            # Ultra-compact single line display
            if len(line_details) <= 20:
                line_numbers = ", ".join(map(str, line_details))
                st.success(f"🔍 **Found on lines:** {line_numbers}")
            else:
                first_few = ", ".join(map(str, line_details[:15]))
                st.success(f"🔍 **Found on {len(line_details)} lines:** {first_few}... (+{len(line_details)-15} more)")
        else:
            st.warning("🔍 No matches found")
    
    # Main document content with line numbers
    with st.expander("📄 Full Document Content (with line numbers)", expanded=True):
        # Create numbered content efficiently
        lines = content.split('\n')
        numbered_lines = [f"{i:4d} | {line}" for i, line in enumerate(lines, 1)]
        numbered_content = '\n'.join(numbered_lines)
        
        st.text_area(
            "Document content with line numbers (use Ctrl+F to search)",
            value=numbered_content,
            height=600,
            help="💡 Use Ctrl+F to search for terms or line numbers (format: '  42 |')",
            key="document_content_viewer"
        )
    
    # Compact search terms display
    if search_terms:
        st.markdown("**🔍 Search with Ctrl+F:**")
        # Display terms in a more compact way
        terms_display = " • ".join([f"`{term}`" for term in search_terms])
        st.markdown(terms_display)
    
    # Compact metrics
    if search_terms and len(search_terms) <= 4:
        cols = st.columns(len(search_terms))
        for i, term in enumerate(search_terms):
            count = content.lower().count(term.lower()) if content and term.strip() else 0
            with cols[i]:
                st.metric(f"'{term[:10]}{'...' if len(term) > 10 else ''}'", f"{count}x")
    
    
    # Show search term summary with better layout
    if search_terms:
        st.markdown("**📊 Search Term Occurrences:**")
        
        # Create metrics in a grid
        num_cols = min(len(search_terms), 4)
        cols = st.columns(num_cols)
        
        for i, term in enumerate(search_terms):
            count = content.lower().count(term.lower())
            with cols[i % num_cols]:
                # Calculate percentage of document that contains this term
                if len(content) > 0:
                    density = (count * len(term) / len(content)) * 100
                    st.metric(
                        label=f"'{term[:15]}{'...' if len(term) > 15 else ''}'",
                        value=f"{count} times",
                        delta=f"{density:.2f}% density"
                    )
                else:
                    st.metric(label=f"'{term}'", value=f"{count} times")

def add_search_term(term: str):
    """Optimized search term addition"""
    if term and (clean_term := term.strip()):
        if clean_term not in st.session_state.search_terms:
            st.session_state.search_terms.append(clean_term)
            st.session_state.input_counter += 1
            # Clear cache when terms change
            if hasattr(st.session_state.search_engine, '_search_single_term'):
                st.session_state.search_engine._search_single_term.cache_clear()
            st.rerun()
        else:
            st.warning(f"'{clean_term}' is already in your search terms!")

def clear_all_terms():
    """Optimized term clearing"""
    st.session_state.update({
        'search_terms': [],
        'current_results': pd.DataFrame(),
        'current_search_key': "",
        'input_counter': st.session_state.input_counter + 1
    })
    # Clear search cache
    if hasattr(st.session_state.search_engine, '_search_single_term'):
        st.session_state.search_engine._search_single_term.cache_clear()
    st.rerun()

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
    
    # Key metrics with updated colors
    col1, col2, col3, col4 = st.columns(4)
    content_lengths = df['Full Contents'].str.len()
    
    metrics = [
        ("📄", "Total Documents", len(df)),
        ("💾", "Total Size (KB)", f"{df.get('Size (KB)', pd.Series()).sum():,.1f}" if 'Size (KB)' in df.columns else "N/A"),
        ("📝", "Total Characters", f"{content_lengths.sum():,}"),
        ("📊", "Avg Doc Length", f"{content_lengths.mean():,.0f} chars")
    ]
    
    for col, (icon, label, value) in zip([col1, col2, col3, col4], metrics):
        # Define color scheme based on metric type
        if "Documents" in label:
            bg_gradient = "linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%)"
            shadow_color = "rgba(30, 58, 138, 0.25)"
        elif "Size" in label or "Total" in label:
            bg_gradient = "linear-gradient(135deg, #059669 0%, #10b981 100%)"
            shadow_color = "rgba(5, 150, 105, 0.25)"
        elif "Characters" in label or "Avg" in label:
            bg_gradient = "linear-gradient(135deg, #dc2626 0%, #ef4444 100%)"
            shadow_color = "rgba(220, 38, 38, 0.25)"
        else:
            bg_gradient = "linear-gradient(135deg, #7c3aed 0%, #8b5cf6 100%)"
            shadow_color = "rgba(124, 58, 237, 0.25)"
            
        with col:
            st.markdown(f"""
            <div style="
                background: {bg_gradient};
                padding: 1rem;
                border-radius: 12px;
                color: white;
                text-align: center;
                margin: 0.5rem 0;
                box-shadow: 0 4px 20px {shadow_color};
                border: 1px solid rgba(255, 255, 255, 0.15);
                transition: transform 0.2s ease;
            " onmouseover="this.style.transform='translateY(-2px)'; this.style.boxShadow='0 6px 25px {shadow_color.replace('0.25', '0.35')}'" 
               onmouseout="this.style.transform='translateY(0px)'; this.style.boxShadow='0 4px 20px {shadow_color}'">
                <h3 style="margin: 0; font-size: 1.5rem;">{icon}</h3>
                <h4 style="margin: 0.5rem 0; font-size: 1.2rem;">{value}</h4>
                <p style="margin: 0; font-size: 0.85rem; opacity: 0.9;">{label}</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Distribution charts with updated colors
    col1, col2 = st.columns(2)
    
    with col1:
        if 'Size (KB)' in df.columns and df['Size (KB)'].notna().any():
            fig = px.histogram(
                x=df['Size (KB)'].dropna(),
                nbins=25, 
                title='📦 Document Size Distribution',
                labels={'x': 'File Size (KB)', 'y': 'Number of Documents'},
                color_discrete_sequence=['#1e3a8a']
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
            color_discrete_sequence=['#3730a3']
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
        
        # Pattern counts using compiled patterns
        pattern_counts = []
        patterns = [
            re.compile(r'<<[^>]*>>', re.IGNORECASE),
            re.compile(r'\{[^}]*\}', re.IGNORECASE),
            re.compile(r'\[[^\]]*\]', re.IGNORECASE)
        ]
        
        for pattern in patterns:
            pattern_counts.append(len(pattern.findall(content)))
        
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
                color_discrete_sequence=['#1e3a8a', '#3730a3']
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
        <p><strong>Enhanced Export:</strong> Includes matched contexts for each document</p>
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
                with st.spinner("Generating Excel export with matched contexts..."):
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

def create_enhanced_export(results_df: pd.DataFrame, search_terms: List[str], search_mode: str) -> bytes:
    """Create enhanced Excel export with matched context lines"""
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
                    'Generated By': ["DocXFilter v3.0"],
                    'Note': ['No documents found matching search criteria']
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Search Summary', index=False)
                
                # Create empty results sheet
                empty_results = pd.DataFrame({'File Name': [], 'Note': []})
                empty_results.to_excel(writer, sheet_name='Search Results', index=False)
            else:
                # Create results sheet with enhanced data including contexts
                export_df = results_df.copy()
                engine = st.session_state.search_engine
                
                # Add match columns and contexts with better error handling
                try:
                    # Add individual term counts (fix the regex parameter issue)
                    for term in search_terms:
                        if 'Full Contents' in export_df.columns:
                            # Use str.count without regex parameter
                            term_lower = term.lower()
                            export_df[f"'{term}' Count"] = export_df['Full Contents'].apply(
                                 lambda x: str(x).lower().count(term_lower) if pd.notna(x) else 0
                            )
                        else:
                            export_df[f"'{term}' Count"] = 0
                    
                    # Add total matches
                    if search_terms and 'Full Contents' in export_df.columns:
                        export_df['Total Matches'] = sum([export_df[f"'{term}' Count"] for term in search_terms])
                    else:
                        export_df['Total Matches'] = 0
                    
                    # Add matched contexts and lines with safe processing
                    matched_contexts = []
                    matched_lines = []
                    
                    for _, row in export_df.iterrows():
                        file_name = row['File Name']
                        
                        try:
                            # Get context lines for this document
                            content = engine._content_cache.get(file_name, '') if hasattr(engine, '_content_cache') else ''
                            
                            if content:
                                # Find lines containing search terms
                                lines = content.split('\n')
                                matching_line_numbers = []
                                matching_line_contents = []
                                
                                for line_num, line in enumerate(lines, 1):
                                    line_matches = []
                                    for term in search_terms:
                                        if term and term.strip() and term.strip().lower() in line.lower():
                                            line_matches.append(term.strip())
                                    
                                    if line_matches:
                                        matching_line_numbers.append(str(line_num))
                                        # Clean the line content for Excel
                                        clean_line = line.strip().replace('\r', '').replace('\n', ' ')[:200]  # Limit length
                                        matching_line_contents.append(f"Line {line_num}: {clean_line}")
                                
                                # Format line numbers
                                line_numbers_str = ", ".join(matching_line_numbers) if matching_line_numbers else "No matches"
                                
                                # Format line contents (limit to first 5 for readability)
                                line_contents_str = " | ".join(matching_line_contents[:5])
                                if len(matching_line_contents) > 5:
                                    line_contents_str += f" | ... and {len(matching_line_contents) - 5} more lines"
                                
                                matched_lines.append(line_numbers_str)
                                matched_contexts.append(line_contents_str if line_contents_str else "No context available")
                            else:
                                matched_lines.append("No content")
                                matched_contexts.append("No content available")
                                
                        except Exception as e:
                            # Fallback for individual document errors
                            matched_lines.append("Error extracting lines")
                            matched_contexts.append(f"Error: {str(e)[:100]}")
                    
                    export_df['Matched Line Numbers'] = matched_lines
                    export_df['Matched Line Contents'] = matched_contexts
                        
                except Exception as e:
                    st.warning(f"Warning: Could not add context data: {str(e)}")
                    # Add empty columns if context extraction fails
                    export_df['Matched Line Numbers'] = "Context extraction failed"
                    export_df['Matched Line Contents'] = "Context extraction failed"
                
                # Determine export columns
                basic_cols = ['File Name']
                optional_cols = ['Size (KB)', 'Content Length (chars)']
                match_cols = [f"'{term}' Count" for term in search_terms if f"'{term}' Count" in export_df.columns]
                
                export_cols = basic_cols + [col for col in optional_cols if col in export_df.columns] + match_cols
                if 'Total Matches' in export_df.columns:
                    export_cols.append('Total Matches')
                if 'Matched Line Numbers' in export_df.columns:
                    export_cols.append('Matched Line Numbers')
                if 'Matched Line Contents' in export_df.columns:
                    export_cols.append('Matched Line Contents')
                
                # Export main results
                export_df[export_cols].to_excel(writer, sheet_name='Search Results', index=False)
                
                # Create summary sheet
                summary_data = {
                    'Search Mode': [search_mode],
                    'Search Terms': [' | '.join(search_terms) if search_terms else 'None'],
                    'Number of Terms': [len(search_terms)],
                    'Results Found': [len(results_df)],
                    'Export Date': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                    'Generated By': ['DocXFilter v3.0'],
                    'Features': ['Search results with matched line numbers and contents included']
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

if __name__ == "__main__":
    main()
