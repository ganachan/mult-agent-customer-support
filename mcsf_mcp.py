import streamlit as st
import asyncio
import httpx
import json
import logging
from typing import Dict, Any, List
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class MicrosoftLearnMCPClient:
    """Official Microsoft Learn MCP Server Client"""
    
    def __init__(self):
        # Official Microsoft Learn MCP endpoint
        self.server_url = "https://learn.microsoft.com/api/mcp"
        self.connected = False
        self.request_id = 1
        self.session_id = None  # For session management
        
    def _get_next_id(self) -> int:
        """Get next request ID"""
        self.request_id += 1
        return self.request_id
    
    async def call_mcp(self, method: str, params: Dict[str, Any] = None) -> Dict[str, Any]:
        """Make an MCP JSON-RPC call using Streamable HTTP transport"""
        payload = {
            "jsonrpc": "2.0",
            "id": self._get_next_id(),
            "method": method
        }
        
        if params:
            payload["params"] = params
        
        # Correct headers for Streamable HTTP transport - Microsoft Learn MCP specific
        headers = {
            "Content-Type": "application/json",
            "Accept": "text/event-stream, application/json",  # Microsoft Learn MCP requires this exact format
            "User-Agent": "Streamlit-MCP-Client/1.0",
            "Cache-Control": "no-cache",
            "Connection": "keep-alive"  # For streaming support
        }
        
        # Add session ID header if we have one
        if hasattr(self, 'session_id') and self.session_id:
            headers["Mcp-Session-Id"] = self.session_id
        
        try:
            async with httpx.AsyncClient(timeout=30.0, follow_redirects=True) as client:
                response = await client.post(self.server_url, json=payload, headers=headers)
                
                # Handle session ID from server
                session_id = response.headers.get("Mcp-Session-Id")
                if session_id:
                    self.session_id = session_id
                
                if response.status_code == 200:
                    # Check content type for response handling
                    content_type = response.headers.get("content-type", "")
                    
                    if "application/json" in content_type:
                        return response.json()
                    elif "text/event-stream" in content_type:
                        # Handle SSE response (streaming)
                        return await self._handle_sse_response(response)
                    else:
                        return response.json()  # Try JSON anyway
                else:
                    return {
                        "error": {
                            "code": response.status_code,
                            "message": f"HTTP {response.status_code}: {response.text[:200]}"
                        }
                    }
                    
        except httpx.TimeoutException:
            return {"error": {"code": -1, "message": "Request timeout - server may be overloaded"}}
        except Exception as e:
            return {"error": {"code": -1, "message": f"Connection error: {str(e)}"}}
    
    async def _handle_sse_response(self, response) -> Dict[str, Any]:
        """Handle Server-Sent Events response for streaming"""
        try:
            # For SSE responses, we need to parse the event stream
            content = await response.aread()
            text = content.decode('utf-8')
            
            # Parse SSE format (simple implementation)
            lines = text.strip().split('\n')
            data_lines = [line[5:] for line in lines if line.startswith('data:')]
            
            if data_lines:
                # Return the last complete JSON object
                for data in reversed(data_lines):
                    try:
                        return json.loads(data)
                    except json.JSONDecodeError:
                        continue
            
            return {"error": {"code": -1, "message": "Could not parse SSE response"}}
            
        except Exception as e:
            return {"error": {"code": -1, "message": f"SSE parsing error: {str(e)}"}}

    
    async def initialize(self) -> tuple[bool, str]:
        """Initialize connection with Microsoft Learn MCP Server"""
        params = {
            "protocolVersion": "2024-11-05",  # Use the supported version
            "capabilities": {
                "tools": {},
                "sampling": {}  # Add sampling capability
            },
            "clientInfo": {
                "name": "Streamlit Microsoft Learn MCP Client",
                "version": "1.0.0"
            }
        }
        
        # Debug: log what we're sending
        st.write("🔍 **Debug Info:**")
        st.write(f"Sending to: {self.server_url}")
        st.write(f"Protocol version: {params['protocolVersion']}")
        
        result = await self.call_mcp("initialize", params)
        
        if "error" not in result and "result" in result:
            self.connected = True
            server_info = result["result"].get("serverInfo", {})
            server_name = server_info.get("name", "Microsoft Learn MCP Server")
            server_version = server_info.get("version", "unknown")
            
            # Show server capabilities
            capabilities = result["result"].get("capabilities", {})
            st.write(f"✅ Server capabilities: {list(capabilities.keys())}")
            
            return True, f"✅ Connected to {server_name} v{server_version}"
        else:
            error = result.get("error", {})
            error_code = error.get("code", "unknown")
            error_msg = error.get("message", "Unknown error")
            
            # Show detailed error for debugging
            st.write(f"❌ Error details:")
            st.write(f"- Code: {error_code}")
            st.write(f"- Message: {error_msg}")
            
            return False, f"❌ Connection failed (Code {error_code}): {error_msg}"
    
    async def list_tools(self) -> List[Dict[str, Any]]:
        """List available tools from Microsoft Learn MCP Server"""
        result = await self.call_mcp("tools/list")
        
        if "error" not in result and "result" in result:
            return result["result"].get("tools", [])
        else:
            return []
    
    async def search_docs(self, query: str) -> Dict[str, Any]:
        """Search Microsoft documentation using microsoft_docs_search"""
        params = {
            "name": "microsoft_docs_search",
            "arguments": {"query": query}
        }
        
        return await self.call_mcp("tools/call", params)
    
    async def test_headers(self) -> Dict[str, Any]:
        """Test what headers the server expects"""
        test_payload = {
            "jsonrpc": "2.0",
            "id": 999,
            "method": "ping"  # Simple test method
        }
        
        # Try different Accept header combinations
        header_combinations = [
            "text/event-stream, application/json",
            "application/json, text/event-stream", 
            "text/event-stream,application/json",
            "application/json,text/event-stream",
            "text/event-stream; application/json",
            "*/*"
        ]
        
        results = {}
        
        for accept_header in header_combinations:
            headers = {
                "Content-Type": "application/json",
                "Accept": accept_header,
                "User-Agent": "Streamlit-MCP-Client/1.0"
            }
            
            try:
                async with httpx.AsyncClient(timeout=10.0) as client:
                    response = await client.post(self.server_url, json=test_payload, headers=headers)
                    results[accept_header] = {
                        "status": response.status_code,
                        "response": response.text[:100] if response.text else "No content"
                    }
            except Exception as e:
                results[accept_header] = {"error": str(e)}
        
        return results

def run_async(coro):
    """Helper to run async functions in Streamlit"""
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        return loop.run_until_complete(coro)
    finally:
        loop.close()

def display_search_results(results: Dict[str, Any]):
    """Display search results from Microsoft Learn MCP"""
    if "error" in results:
        error_info = results["error"]
        st.error(f"❌ Search failed: {error_info.get('message', 'Unknown error')}")
        return
    
    if "result" not in results:
        st.warning("⚠️ No results structure found")
        st.json(results)
        return
    
    content = results["result"].get("content", [])
    
    if not content:
        st.info("ℹ️ No results found for your query")
        return
    
    st.success(f"✅ Found {len(content)} results from Microsoft Learn")
    
    for i, item in enumerate(content):
        if isinstance(item, dict) and item.get("type") == "text":
            # This is a text content block
            text_content = item.get("text", "")
            
            # Try to extract title and content
            lines = text_content.split('\n')
            title = lines[0] if lines else f"Result {i+1}"
            
            with st.expander(f"📄 {title[:100]}{'...' if len(title) > 100 else ''}", expanded=i==0):
                st.markdown(text_content)
        else:
            # Handle other content types
            with st.expander(f"📄 Result {i+1}", expanded=i==0):
                if isinstance(item, str):
                    st.markdown(item)
                else:
                    st.json(item)

def display_doc_content(results: Dict[str, Any]):
    """Display fetched document content"""
    if "error" in results:
        error_info = results["error"]
        st.error(f"❌ Failed to fetch document: {error_info.get('message', 'Unknown error')}")
        return
    
    if "result" not in results:
        st.warning("⚠️ No document content found")
        st.json(results)
        return
    
    content = results["result"].get("content", [])
    
    if not content:
        st.info("ℹ️ Document appears to be empty")
        return
    
    st.success("✅ Document fetched successfully")
    
    # Display document content
    for item in content:
        if isinstance(item, dict) and item.get("type") == "text":
            st.markdown(item.get("text", ""))
        elif isinstance(item, str):
            st.markdown(item)
        else:
            st.json(item)

def main():
    st.set_page_config(
        page_title="Microsoft Learn MCP Explorer",
        page_icon="📚",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Header with official branding
    st.markdown("""
    <div style="background: linear-gradient(90deg, #0078d4 0%, #106ebe 100%); color: white; padding: 1.5rem; border-radius: 10px; margin-bottom: 1rem;">
        <h1>📚 Microsoft Learn MCP Explorer</h1>
        <p>Official integration with Microsoft Learn MCP Server</p>
        <small>Endpoint: https://learn.microsoft.com/api/mcp</small>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    if 'connected' not in st.session_state:
        st.session_state.connected = False
    if 'client' not in st.session_state:
        st.session_state.client = None
    if 'tools' not in st.session_state:
        st.session_state.tools = []
    
    # Sidebar
    with st.sidebar:
        st.header("🔌 Connection")
        
        # Server info
        st.info("""
        **Official Microsoft Learn MCP Server**
        - 🌐 Endpoint: learn.microsoft.com/api/mcp
        - 🔓 No authentication required  
        - 🚀 Streamable HTTP transport (2025-03-26)
        - 📊 Real-time Microsoft documentation
        """)
        
        # Connection status
        if st.session_state.connected:
            st.success("🟢 Connected to Microsoft Learn MCP")
            if hasattr(st.session_state.client, 'session_id') and st.session_state.client.session_id:
                st.info(f"Session ID: {st.session_state.client.session_id[:12]}...")
        else:
            st.error("🔴 Not Connected")
        
        # Connection button
        if not st.session_state.connected:
            if st.button("🔌 Connect to Microsoft Learn MCP", type="primary", use_container_width=True):
                with st.spinner("Connecting to Microsoft Learn MCP Server..."):
                    client = MicrosoftLearnMCPClient()
                    success, message = run_async(client.initialize())
                    
                    if success:
                        st.session_state.client = client
                        st.session_state.connected = True
                        st.success(message)
                        
                        # Load available tools
                        with st.spinner("Loading available tools..."):
                            tools = run_async(client.list_tools())
                            st.session_state.tools = tools
                            
                            if tools:
                                st.info(f"📋 Loaded {len(tools)} tools")
                                
                                # Show tool names briefly
                                tool_names = [tool.get('name', 'Unknown') for tool in tools]
                                st.success(f"Available tools: {', '.join(tool_names)}")
                            else:
                                st.warning("⚠️ No tools found - this might be a protocol issue")
                                st.info("💡 The server is connected, but tool discovery may need adjustment")
                    else:
                        st.error(message)
                        st.info("💡 Try refreshing the page and connecting again")
        else:
            st.success("🟢 Connected to Microsoft Learn MCP")
            
            # Show available tools
            if st.session_state.tools:
                with st.expander("🛠️ Available Tools"):
                    for tool in st.session_state.tools:
                        st.markdown(f"**{tool.get('name', 'Unknown')}**")
                        st.markdown(f"_{tool.get('description', 'No description')}_")
                        
                        # Show input schema if available
                        if 'inputSchema' in tool:
                            with st.expander(f"Schema for {tool['name']}", expanded=False):
                                st.json(tool['inputSchema'])
            
            # Disconnect button
            if st.button("🔄 Disconnect", type="secondary", use_container_width=True):
                st.session_state.connected = False
                st.session_state.client = None
                st.session_state.tools = []
                st.rerun()
    
    # Main content
    if st.session_state.connected and st.session_state.client:
        # Tab layout for different functions
        tab1, tab2 = st.tabs(["🔍 Search Documentation", "📖 Fetch Specific Document"])
        
        with tab1:
            st.header("🔍 Search Microsoft Documentation")
            st.markdown("Use semantic search to find relevant Microsoft Learn content")
            
            # Search form
            with st.form("search_form"):
                query = st.text_area(
                    "Search Query:",
                    height=100,
                    placeholder="e.g., 'How to create Azure Functions with Python'\n'ASP.NET Core authentication setup'\n'Entity Framework Core migrations'",
                    help="Enter your search query. The MCP server uses semantic search to find the most relevant Microsoft documentation."
                )
                
                search_button = st.form_submit_button("🔍 Search Microsoft Learn", type="primary", use_container_width=True)
            
            if search_button and query.strip():
                # Add to search history
                if 'search_history' not in st.session_state:
                    st.session_state.search_history = []
                
                if query not in st.session_state.search_history:
                    st.session_state.search_history.insert(0, query)
                    st.session_state.search_history = st.session_state.search_history[:10]  # Keep last 10
                
                # Perform search
                with st.spinner("🔍 Searching Microsoft Learn documentation..."):
                    results = run_async(st.session_state.client.search_docs(query))
                    display_search_results(results)
            
            # Search history
            if st.session_state.get('search_history'):
                st.markdown("### 📋 Recent Searches")
                cols = st.columns(min(3, len(st.session_state.search_history)))
                
                for i, recent_query in enumerate(st.session_state.search_history[:3]):
                    with cols[i]:
                        if st.button(f"🔄 {recent_query[:25]}{'...' if len(recent_query) > 25 else ''}", key=f"recent_{i}"):
                            st.session_state.repeat_query = recent_query
                            st.rerun()
        
        with tab2:
            st.header("📖 Fetch Specific Document")
            st.markdown("Retrieve and convert a specific Microsoft Learn page to markdown")
            
            # Fetch form
            with st.form("fetch_form"):
                doc_url = st.text_input(
                    "Microsoft Learn URL:",
                    placeholder="https://learn.microsoft.com/en-us/azure/...",
                    help="Enter the full URL of a Microsoft Learn documentation page"
                )
                
                st.markdown("**Example URLs:**")
                st.markdown("""
                - `https://learn.microsoft.com/en-us/azure/azure-functions/functions-overview`
                - `https://learn.microsoft.com/en-us/aspnet/core/tutorials/first-mvc-app/`
                - `https://learn.microsoft.com/en-us/dotnet/core/tutorials/with-visual-studio`
                """)
                
                fetch_button = st.form_submit_button("📖 Fetch Document", type="primary", use_container_width=True)
            
            if fetch_button and doc_url.strip():
                # Validate URL
                if not doc_url.startswith("https://learn.microsoft.com"):
                    st.error("❌ Please enter a valid Microsoft Learn URL (must start with https://learn.microsoft.com)")
                else:
                    with st.spinner("📖 Fetching document from Microsoft Learn..."):
                        results = run_async(st.session_state.client.fetch_doc(doc_url))
                        display_doc_content(results)
        
        # Example queries section
        st.markdown("---")
        st.markdown("### 💡 Example Queries to Try")
        
        examples = [
            "Azure CLI commands to create an Azure Container App with managed identity",
            "How to implement IHttpClientFactory in .NET 8 minimal API",
            "Complete guide for implementing authentication in ASP.NET Core",
            "Step-by-step tutorial for deploying .NET application to Azure App Service",
            "Azure Functions end-to-end development guide",
            "Entity Framework Core migration best practices",
            "Power BI REST API integration examples"
        ]
        
        cols = st.columns(2)
        for i, example in enumerate(examples):
            with cols[i % 2]:
                if st.button(f"💡 {example}", key=f"example_{i}"):
                    st.session_state.example_query = example
                    # Trigger search with example
                    with st.spinner("🔍 Searching..."):
                        results = run_async(st.session_state.client.search_docs(example))
                        st.markdown(f"### Results for: {example}")
                        display_search_results(results)
    
    else:
        # Welcome screen
        st.info("👈 Please connect to the Microsoft Learn MCP Server using the sidebar")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 🚀 What is Microsoft Learn MCP?")
            st.markdown("""
            The **Microsoft Learn MCP Server** is an official remote server that provides:
            
            - **🔍 Semantic Search**: Advanced search through Microsoft's official documentation
            - **📄 Document Fetching**: Retrieve and convert specific pages to markdown
            - **🔄 Real-time Updates**: Access the latest Microsoft documentation as it's published
            - **🏢 Official Content**: Trusted, authoritative Microsoft technical content
            """)
        
        with col2:
            st.markdown("### 🛠️ Available Tools")
            st.markdown("""
            **microsoft_docs_search**
            - Performs semantic search against Microsoft official technical documentation
            - Input: `query` (string) - Your search query
            
            **microsoft_docs_fetch**  
            - Fetch and convert a Microsoft documentation page into markdown format
            - Input: `url` (string) - URL of the documentation page to read
            """)
        
        # Getting started
        st.markdown("### 🎯 Getting Started")
        st.markdown("""
        1. **Click 'Connect to Microsoft Learn MCP'** in the sidebar
        2. **Wait for connection confirmation** and tool loading
        3. **Use the Search tab** to find relevant documentation
        4. **Use the Fetch tab** to retrieve specific documents
        5. **Try the example queries** to see the system in action
        """)

if __name__ == "__main__":
    main()
