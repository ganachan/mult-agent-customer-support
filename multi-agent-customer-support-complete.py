"""
Multi-Agent Customer Support System with Azure AI Search RAG Integration, Avatar, Email Notifications, and Microsoft Learn MCP Search
Stores cases in memory, retrieves relevant knowledge from Azure AI Search, generates avatar videos, sends email notifications, and provides MCP search for additional support

Required Environment Variables:
- AZURE_OPENAI_ENDPOINT
- AZURE_OPENAI_API_KEY  
- AZURE_OPENAI_DEPLOYMENT_NAME
- AZURE_SEARCH_ENDPOINT
- AZURE_SEARCH_KEY
- AZURE_SEARCH_INDEX
- AZCOSMOS_CONNSTR
- AZCOSMOS_DATABASE_NAME
- AZCOSMOS_CONTAINER_NAME
- SPEECH_SUBSCRIPTION_KEY
- SPEECH_ENDPOINT
- BLOB_CONNECTION_STRING
- BLOB_CONTAINER_NAME
- AZURE_COMMUNICATION_EMAIL_CONNECTION_STRING
- EMAIL_SENDER_ADDRESS

Installation:
pip install streamlit python-dotenv requests openai semantic-kernel azure-search-documents pymongo azure-storage-blob azure-communication-email httpx
"""

import streamlit as st
import os
import json
import asyncio
import time
import uuid
import requests
import httpx
from datetime import datetime, timedelta, timezone
from io import StringIO
from dotenv import load_dotenv
from openai import AzureOpenAI
from typing import Dict, Any, List

# Azure imports
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions

# Cosmos DB imports
from pymongo import MongoClient
from pymongo.errors import ConnectionFailure, OperationFailure

# Semantic Kernel imports
from semantic_kernel.agents import AgentGroupChat, ChatCompletionAgent
from semantic_kernel.connectors.ai.open_ai.services.azure_chat_completion import AzureChatCompletion
from semantic_kernel.contents.chat_message_content import ChatMessageContent
from semantic_kernel.contents.utils.author_role import AuthorRole
from semantic_kernel import Kernel

# Email imports
from azure.communication.email import EmailClient

# Load environment variables
load_dotenv(override=True)

# Page configuration
st.set_page_config(
    layout="wide",
    page_title="Microsoft Customer Support AI Agent",
    page_icon="🤖",
    initial_sidebar_state="expanded"
)

# Microsoft-style CSS
st.markdown("""
<style>
    .stApp {
        background-color: #F3F2F1;
        background-image: linear-gradient(to bottom right, #F3F2F1, #E1DFDD);
    }
    
    .main-header {
        background: linear-gradient(90deg, #0078D4 0%, #106EBE 100%);
        color: white;
        padding: 20px;
        border-radius: 8px;
        margin-bottom: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    .status-card {
        background: white;
        padding: 15px;
        border-radius: 6px;
        border-left: 4px solid #0078D4;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        margin: 10px 0;
    }
    
    .success-card {
        background: #DFF6DD;
        border-left-color: #107C10;
        color: #107C10;
    }
    
    .warning-card {
        background: #FFF4CE;
        border-left-color: #FF8C00;
        color: #8A6914;
    }
    
    .info-card {
        background: #CCE7FF;
        border-left-color: #0078D4;
        color: #004578;
    }
    
    .mcp-card {
        background: #F0F8FF;
        border-left-color: #4A90E2;
        border: 1px solid #E0E8F0;
    }
    
    .stTextInput > div > div > input {
        background-color: white;
        border: 1px solid #D1D1D1;
        border-radius: 4px;
    }
    
    .stTextArea > div > div > textarea {
        background-color: white;
        border: 1px solid #D1D1D1;
        border-radius: 4px;
    }
    
    .stButton > button {
        background-color: #0078D4;
        color: white;
        border: none;
        border-radius: 4px;
        font-weight: 600;
        transition: background-color 0.2s;
    }
    
    .stButton > button:hover {
        background-color: #106EBE;
    }
    
    .agent-section {
        background: white;
        padding: 20px;
        border-radius: 8px;
        margin: 15px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
        border-top: 3px solid #0078D4;
    }
    
    .mcp-section {
        background: white;
        padding: 20px;
        border-radius: 8px;
        margin: 15px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
        border-top: 3px solid #4A90E2;
        border: 1px solid #E0E8F0;
    }
    
    h1, h2, h3 {
        color: #323130;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    .microsoft-logo {
        font-size: 24px;
        font-weight: bold;
        color: #0078D4;
    }
    
    .search-result {
        background: #f8f9fa;
        padding: 15px;
        border-radius: 6px;
        margin: 10px 0;
        border-left: 3px solid #4A90E2;
    }
</style>
""", unsafe_allow_html=True)

class MicrosoftLearnMCPClient:
    """Microsoft Learn MCP Server Client for additional support queries"""
    
    def __init__(self):
        self.server_url = "https://learn.microsoft.com/api/mcp"
        self.connected = False
        self.request_id = 1
        self.session_id = None
        
    def _get_next_id(self) -> int:
        self.request_id += 1
        return self.request_id
    
    async def call_mcp(self, method: str, params: Dict[str, Any] = None) -> Dict[str, Any]:
        """Make an MCP JSON-RPC call"""
        payload = {
            "jsonrpc": "2.0",
            "id": self._get_next_id(),
            "method": method
        }
        
        if params:
            payload["params"] = params
        
        headers = {
            "Content-Type": "application/json",
            "Accept": "text/event-stream, application/json",
            "User-Agent": "Microsoft-Support-MCP-Client/1.0",
            "Cache-Control": "no-cache",
            "Connection": "keep-alive"
        }
        
        if hasattr(self, 'session_id') and self.session_id:
            headers["Mcp-Session-Id"] = self.session_id
        
        try:
            async with httpx.AsyncClient(timeout=30.0, follow_redirects=True) as client:
                response = await client.post(self.server_url, json=payload, headers=headers)
                
                session_id = response.headers.get("Mcp-Session-Id")
                if session_id:
                    self.session_id = session_id
                
                if response.status_code == 200:
                    content_type = response.headers.get("content-type", "")
                    
                    if "application/json" in content_type:
                        return response.json()
                    elif "text/event-stream" in content_type:
                        return await self._handle_sse_response(response)
                    else:
                        return response.json()
                else:
                    return {
                        "error": {
                            "code": response.status_code,
                            "message": f"HTTP {response.status_code}: {response.text[:200]}"
                        }
                    }
                    
        except httpx.TimeoutException:
            return {"error": {"code": -1, "message": "Request timeout"}}
        except Exception as e:
            return {"error": {"code": -1, "message": f"Connection error: {str(e)}"}}
    
    async def _handle_sse_response(self, response) -> Dict[str, Any]:
        """Handle Server-Sent Events response"""
        try:
            content = await response.aread()
            text = content.decode('utf-8')
            
            lines = text.strip().split('\n')
            data_lines = [line[5:] for line in lines if line.startswith('data:')]
            
            if data_lines:
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
            "protocolVersion": "2024-11-05",
            "capabilities": {
                "tools": {},
                "sampling": {}
            },
            "clientInfo": {
                "name": "Microsoft Support MCP Client",
                "version": "1.0.0"
            }
        }
        
        result = await self.call_mcp("initialize", params)
        
        if "error" not in result and "result" in result:
            self.connected = True
            server_info = result["result"].get("serverInfo", {})
            server_name = server_info.get("name", "Microsoft Learn MCP Server")
            server_version = server_info.get("version", "unknown")
            return True, f"Connected to {server_name} v{server_version}"
        else:
            error = result.get("error", {})
            error_msg = error.get("message", "Unknown error")
            return False, f"Connection failed: {error_msg}"
    
    async def list_tools(self) -> List[Dict[str, Any]]:
        """List available tools"""
        result = await self.call_mcp("tools/list")
        
        if "error" not in result and "result" in result:
            return result["result"].get("tools", [])
        else:
            return []
    
    async def search_docs(self, query: str) -> Dict[str, Any]:
        """Search Microsoft documentation"""
        params = {
            "name": "microsoft_docs_search",
            "arguments": {"query": query}
        }
        
        return await self.call_mcp("tools/call", params)

class Config:
    """Application configuration"""
    def __init__(self):
        # Azure OpenAI
        self.azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
        self.api_key = os.getenv("AZURE_OPENAI_API_KEY")
        self.deployment_name = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME")
        self.api_version = "2024-02-01"
        
        # Azure AI Search
        self.search_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT")
        self.search_key = os.getenv("AZURE_SEARCH_KEY")
        self.search_index = os.getenv("AZURE_SEARCH_INDEX", "case1rag")
        
        # Cosmos DB
        self.cosmos_connection_string = os.getenv("AZCOSMOS_CONNSTR")
        self.cosmos_database_name = os.getenv("AZCOSMOS_DATABASE_NAME", "customerservice")
        self.cosmos_container_name = os.getenv("AZCOSMOS_CONTAINER_NAME", "cases")
        
        # Avatar/Speech
        self.speech_endpoint = os.getenv("SPEECH_ENDPOINT")
        self.speech_key = os.getenv("SPEECH_SUBSCRIPTION_KEY")
        self.blob_connection_string = os.getenv("BLOB_CONNECTION_STRING")
        self.blob_container_name = os.getenv("BLOB_CONTAINER_NAME", "avatar-videos")
        self.api_version_speech = "2024-08-01"
        
        # Email Service
        self.email_connection_string = os.getenv("AZURE_COMMUNICATION_EMAIL_CONNECTION_STRING")
        self.email_sender_address = os.getenv("EMAIL_SENDER_ADDRESS", "DoNotReply@40a3f6fd-3da7-481e-b3fc-7b20dd6c32c2.azurecomm.net")
        
    def validate(self):
        required = [self.azure_endpoint, self.api_key, self.deployment_name]
        return all(required)
    
    def validate_search(self):
        return all([self.search_endpoint, self.search_key, self.search_index])
    
    def validate_cosmos(self):
        return bool(self.cosmos_connection_string)
    
    def validate_avatar(self):
        return all([self.speech_endpoint, self.speech_key, self.blob_connection_string])
    
    def validate_email(self):
        return bool(self.email_connection_string)

class AvatarService:
    """Handles avatar video generation and management"""
    
    def __init__(self, config, show_status=True):
        self.config = config
        self.connected = False
        self.blob_service_client = None
        self.container_client = None
        self.show_status = show_status
        
        if config.validate_avatar():
            try:
                self.blob_service_client = BlobServiceClient.from_connection_string(config.blob_connection_string)
                self.container_client = self.blob_service_client.get_container_client(config.blob_container_name)
                self.connected = True
                self._create_container_if_not_exists()
            except Exception as e:
                self.connected = False
    
    def _create_container_if_not_exists(self):
        """Create blob container if it doesn't exist"""
        try:
            self.container_client.create_container()
        except Exception:
            pass  # Container already exists
    
    def _authenticate(self):
        """Get authentication headers for speech service"""
        return {'Ocp-Apim-Subscription-Key': self.config.speech_key}
    
    def _create_job_id(self):
        """Generate unique job ID"""
        return str(uuid.uuid4())
    
    def check_available_avatars(self):
        """Check available avatar characters"""
        if not self.connected:
            return None
        
        url = f'{self.config.speech_endpoint}/avatar/characters?api-version={self.config.api_version_speech}'
        headers = self._authenticate()
        
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                return response.json()
            else:
                st.info(f"Could not fetch available avatars: {response.status_code}")
                return None
        except Exception as e:
            st.info(f"Could not check available avatars: {e}")
            return None
    
    def submit_avatar_synthesis(self, job_id: str, text: str, customer_name: str = "Valued Customer", avatar_character: str = "ava"):
        """Submit text-to-speech avatar synthesis job"""
        if not self.connected:
            return None
        
        url = f'{self.config.speech_endpoint}/avatar/batchsyntheses/{job_id}?api-version={self.config.api_version_speech}'
        headers = {'Content-Type': 'application/json'}
        headers.update(self._authenticate())
        
        avatar_name = "Sara" if avatar_character == "sara" else "Lisa"
        personalized_text = f"""
        Hello {customer_name}, I'm {avatar_name} from Microsoft Support.
        
        Your support case has been resolved by our AI team.
        
        {text}
        
        Thank you for choosing Microsoft.
        """
        
        payload = {
            'synthesisConfig': {
                "voice": "en-US-AvaMultilingualNeural",
                "outputFormat": "riff-24khz-16bit-mono-pcm"
            },
            "inputKind": "plainText",
            "inputs": [{"content": personalized_text.strip()}],
            "avatarConfig": {
                "customized": False,
                "talkingAvatarCharacter": "lisa",
                "talkingAvatarStyle": "graceful-sitting",
                "subtitleType": "hard_embedded",
                "videoFormat": "mp4",
                "videoCodec": "h264",
                "subtitleType": "soft_embedded",
                "backgroundColor": "#FFFFFF"
            }
        }
        
        try:
            response = requests.put(url, json=payload, headers=headers)
            if response.status_code < 400:
                return response.json().get("id")
            else:
                error_msg = f"Avatar synthesis error: {response.status_code} - {response.text}"
                st.error(error_msg)
                return None
        except Exception as e:
            st.error(f"Failed to submit avatar job: {e}")
            return None
    
    def get_synthesis_status(self, job_id):
        """Check synthesis job status"""
        if not self.connected:
            return None, None
        
        url = f'{self.config.speech_endpoint}/avatar/batchsyntheses/{job_id}?api-version={self.config.api_version_speech}'
        headers = self._authenticate()
        
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()
            
            status = data.get('status', 'Unknown')
            if status == 'Failed':
                error_details = data.get('properties', {}).get('error', {})
                st.error(f"Avatar synthesis failed: {error_details}")
                return None, data
            elif status == 'Succeeded':
                return data.get('outputs', {}).get('result'), data
            else:
                return None, data
                
        except Exception as e:
            st.error(f"Failed to get synthesis status: {e}")
            return None, None
    
    def upload_video_to_blob(self, video_data, blob_name):
        """Upload video to blob storage"""
        if not self.connected:
            return None
        
        try:
            blob_client = self.container_client.get_blob_client(blob_name)
            blob_client.upload_blob(video_data, overwrite=True)
            return blob_client.url
        except Exception as e:
            st.error(f"Failed to upload video: {e}")
            return None
    
    def generate_sas_token(self, blob_name):
        """Generate SAS token for secure video access"""
        if not self.connected:
            return None
        
        try:
            sas_token = generate_blob_sas(
                account_name=self.blob_service_client.account_name,
                container_name=self.config.blob_container_name,
                blob_name=blob_name,
                account_key=self.blob_service_client.credential.account_key,
                permission=BlobSasPermissions(read=True),
                expiry=datetime.now(timezone.utc) + timedelta(hours=24)
            )
            return sas_token
        except Exception as e:
            st.error(f"Failed to generate SAS token: {e}")
            return None

class EmailService:
    """Handles email notifications for case completion"""
    
    def __init__(self, config, show_status=True):
        self.config = config
        self.connected = False
        self.email_client = None
        self.show_status = show_status
        
        if config.validate_email():
            try:
                self.email_client = EmailClient.from_connection_string(config.email_connection_string)
                self.connected = True
            except Exception as e:
                self.connected = False
    
    def send_case_notification(self, case_data, resolution_summary, recipient_email, manager_name="Support Manager"):
        """Send email notification when case is completed"""
        if not self.connected:
            return False, "Email service not connected"
        
        try:
            case_number = case_data.get('Case Number', 'Unknown')
            customer_name = case_data.get('Customer Name', 'Unknown Customer')
            subject = f"Case #{case_number} Resolved - {customer_name}"
            
            plain_text = self._create_plain_text_email(case_data, resolution_summary, manager_name)
            html_content = self._create_html_email(case_data, resolution_summary, manager_name)
            
            message = {
                "senderAddress": self.config.email_sender_address,
                "recipients": {
                    "to": [{"address": recipient_email}]
                },
                "content": {
                    "subject": subject,
                    "plainText": plain_text,
                    "html": html_content
                }
            }
            
            poller = self.email_client.begin_send(message)
            result = poller.result()

            if hasattr(result, 'message_id'):
                message_id = result.message_id
            elif isinstance(result, dict) and 'messageId' in result:
                message_id = result['messageId']
            elif isinstance(result, dict) and 'id' in result:
                message_id = result['id']
            else:
                message_id = "Unknown"
            return True, f"Email sent successfully. Message ID: {message_id}"
        except Exception as e:
            return False, f"Failed to send email: {str(e)}"
    
    def _create_plain_text_email(self, case_data, resolution_summary, manager_name):
        """Create plain text version of the email"""
        return f"""Hello {manager_name},

A customer support case has been successfully resolved by our AI agent system.

CASE DETAILS:
- Case Number: {case_data.get('Case Number', 'N/A')}
- Customer: {case_data.get('Customer Name', 'N/A')}
- Organization: {case_data.get('Organization', 'N/A')}
- Issue: {case_data.get('Issue Description', 'N/A')}
- Duration: {case_data.get('Issue Duration', 'N/A')}
- Root Cause: {case_data.get('Root Cause', 'N/A')}

RESOLUTION SUMMARY:
{resolution_summary}

This case was processed and resolved automatically by our multi-agent AI system with RAG integration.

Best regards,
Microsoft AI Customer Support System
"""
    
    def _create_html_email(self, case_data, resolution_summary, manager_name):
        """Create HTML version of the email"""
        return f"""
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
        .header {{ background: linear-gradient(90deg, #0078D4 0%, #106EBE 100%); color: white; padding: 20px; border-radius: 8px; }}
        .case-details {{ background: #f8f9fa; padding: 15px; border-left: 4px solid #0078D4; margin: 20px 0; }}
        .resolution {{ background: #DFF6DD; padding: 15px; border-left: 4px solid #107C10; margin: 20px 0; }}
        .footer {{ color: #666; font-size: 12px; margin-top: 30px; }}
        .detail-row {{ margin: 8px 0; }}
        .label {{ font-weight: bold; color: #323130; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Microsoft AI Customer Support</h1>
        <p>Case Resolution Notification</p>
    </div>
    
    <p>Hello <strong>{manager_name}</strong>,</p>
    
    <p>A customer support case has been successfully resolved by our AI agent system.</p>
    
    <div class="case-details">
        <h3>Case Details</h3>
        <div class="detail-row"><span class="label">Case Number:</span> {case_data.get('Case Number', 'N/A')}</div>
        <div class="detail-row"><span class="label">Customer:</span> {case_data.get('Customer Name', 'N/A')}</div>
        <div class="detail-row"><span class="label">Organization:</span> {case_data.get('Organization', 'N/A')}</div>
        <div class="detail-row"><span class="label">Issue:</span> {case_data.get('Issue Description', 'N/A')}</div>
        <div class="detail-row"><span class="label">Duration:</span> {case_data.get('Issue Duration', 'N/A')}</div>
        <div class="detail-row"><span class="label">Root Cause:</span> {case_data.get('Root Cause', 'N/A')}</div>
    </div>
    
    <div class="resolution">
        <h3>Resolution Summary</h3>
        <p>{resolution_summary.replace(chr(10), '<br>')}</p>
    </div>
    
    <p>This case was processed and resolved automatically by our multi-agent AI system with RAG integration.</p>
    
    <div class="footer">
        <p>Best regards,<br>
        Microsoft AI Customer Support System<br>
        <em>Powered by Azure AI Services</em></p>
    </div>
</body>
</html>
"""

class CosmosDBService:
    """Handles Cosmos DB operations for case management"""
    
    def __init__(self, config, show_status=True):
        self.config = config
        self.client = None
        self.database = None
        self.collection = None
        self.connected = False
        self.show_status = show_status
        
        if config.validate_cosmos():
            self._connect()
    
    def _connect(self):
        """Initialize connection to Cosmos DB"""
        try:
            self.client = MongoClient(self.config.cosmos_connection_string)
            self.client.admin.command('ping')
            self.database = self.client[self.config.cosmos_database_name]
            self.collection = self.database[self.config.cosmos_container_name]
            self.connected = True
        except Exception as e:
            self.connected = False
    
    def save_case(self, case_data):
        """Save a new case to Cosmos DB"""
        if not self.connected:
            return None
        
        try:
            case_data_copy = case_data.copy()
            if '_id' in case_data_copy:
                del case_data_copy['_id']
            
            case_data_copy.update({
                "created_at": datetime.utcnow().isoformat(),
                "updated_at": datetime.utcnow().isoformat(),
                "status": "created",
                "processing_log": [],
                "resolution_summary": ""
            })
            
            result = self.collection.insert_one(case_data_copy)
            case_id = str(result.inserted_id)
            return case_id
        except Exception as e:
            st.error(f"Failed to save case: {e}")
            return None
    
    def log_agent_action(self, case_id, agent_name, action_type, details):
        """Log agent actions"""
        if not self.connected or not case_id:
            return False
        
        try:
            log_entry = {
                "timestamp": datetime.utcnow().isoformat(),
                "agent": agent_name,
                "action_type": action_type,
                "details": details
            }
            
            result = self.collection.update_one(
                {"_id": case_id},
                {
                    "$push": {"processing_log": log_entry},
                    "$set": {
                        "updated_at": datetime.utcnow().isoformat(),
                        "current_agent": agent_name,
                        "status": f"processing_{agent_name.lower()}"
                    }
                }
            )
            return result.modified_count > 0
        except Exception as e:
            st.error(f"Failed to log action: {e}")
            return False
    
    def complete_case(self, case_id, resolution_summary):
        """Mark case as completed with resolution summary"""
        if not self.connected or not case_id:
            return False
        
        try:
            result = self.collection.update_one(
                {"_id": case_id},
                {
                    "$set": {
                        "updated_at": datetime.utcnow().isoformat(),
                        "completed_at": datetime.utcnow().isoformat(),
                        "status": "completed",
                        "resolution_summary": resolution_summary
                    }
                }
            )
            return result.modified_count > 0
        except Exception as e:
            st.error(f"Failed to complete case: {e}")
            return False
    
    def get_case_resolution(self, case_id):
        """Get case resolution summary"""
        if not self.connected or not case_id:
            return None
        
        try:
            case = self.collection.find_one(
                {"_id": case_id}, 
                {"resolution_summary": 1, "Customer Name": 1, "Case Number": 1}
            )
            return case
        except Exception as e:
            st.error(f"Failed to get case resolution: {e}")
            return None

class InMemoryStorage:
    """Stores cases in memory with Cosmos DB integration"""
    def __init__(self, cosmos_service=None):
        if 'cases' not in st.session_state:
            st.session_state.cases = []
        if 'current_case_id' not in st.session_state:
            st.session_state.current_case_id = None
        if 'resolution_summary' not in st.session_state:
            st.session_state.resolution_summary = ""
        if 'last_saved_case_hash' not in st.session_state:
            st.session_state.last_saved_case_hash = None
        self.cosmos_service = cosmos_service
    
    def _get_case_hash(self, case_data):
        """Generate hash for case to prevent duplicates"""
        import hashlib
        key_fields = f"{case_data.get('Case Number', '')}{case_data.get('Customer Name', '')}{case_data.get('Issue Description', '')}"
        return hashlib.md5(key_fields.encode()).hexdigest()
    
    def save_case(self, case_data):
        try:
            case_hash = self._get_case_hash(case_data)
            if st.session_state.last_saved_case_hash == case_hash:
                st.warning("This case has already been saved. Skipping duplicate.")
                return True
            
            case_data['memory_id'] = f"case_{len(st.session_state.cases) + 1}"
            case_data['timestamp'] = datetime.now().isoformat()
            st.session_state.cases.append(case_data)
            
            if self.cosmos_service and self.cosmos_service.connected:
                cosmos_id = self.cosmos_service.save_case(case_data.copy())
                if cosmos_id:
                    st.session_state.current_case_id = cosmos_id
            
            st.session_state.last_saved_case_hash = case_hash
            
            return True
        except Exception as e:
            st.error(f"Error saving case: {e}")
            return False
    
    def fetch_latest_case(self):
        if st.session_state.cases:
            return st.session_state.cases[-1]
        return None

class KnowledgeService:
    """Handles knowledge retrieval from Azure AI Search"""
    def __init__(self, config, show_status=True):
        self.config = config
        self.search_client = None
        self.embedding_client = None
        self.show_status = show_status
        
        if config.validate_search():
            try:
                self.search_client = SearchClient(
                    endpoint=config.search_endpoint,
                    index_name=config.search_index,
                    credential=AzureKeyCredential(config.search_key)
                )
                
                self.embedding_client = AzureOpenAI(
                    api_key=config.api_key,
                    api_version=config.api_version,
                    azure_endpoint=config.azure_endpoint
                )
            except Exception as e:
                pass  # Silently handle connection issues
    
    def get_embedding(self, text):
        """Generate embedding for vector search"""
        if not self.embedding_client:
            return None
        
        try:
            response = self.embedding_client.embeddings.create(
                input=text,
                model="text-embedding-ada-002"
            )
            return response.data[0].embedding
        except Exception as e:
            st.error(f"Embedding generation failed: {e}")
            return None
    
    def search_similar_cases(self, issue_description, top_k=3):
        """Search for similar cases using vector similarity"""
        if not self.search_client:
            return []
        
        try:
            query_embedding = self.get_embedding(issue_description)
            
            if query_embedding:
                results = self.search_client.search(
                    search_text="",
                    vector_queries=[{
                        "kind": "vector",
                        "vector": query_embedding,
                        "fields": "text_vector",
                        "k": top_k
                    }],
                    select=["chunk_id", "parent_id", "chunk", "title"],
                    top=top_k
                )
            else:
                results = self.search_client.search(
                    search_text=issue_description,
                    select=["chunk_id", "parent_id", "chunk", "title"],
                    top=top_k
                )
            
            return list(results)
        except Exception as e:
            st.error(f"Search error: {e}")
            return []

class AIService:
    """Handles Azure OpenAI operations"""
    def __init__(self, config):
        try:
            self.client = AzureOpenAI(
                api_key=config.api_key,
                api_version=config.api_version,
                azure_endpoint=config.azure_endpoint
            )
            self.deployment_name = config.deployment_name
        except Exception as e:
            st.error(f"Azure OpenAI initialization error: {e}")
            self.client = None
    
    def extract_labels_from_transcript(self, transcript):
        if not self.client:
            return {}
        
        prompt = f"""Extract customer support information from this transcript and return ONLY valid JSON:

{{
  "Organization": "Company name",
  "Case Number": "Case/ticket number if mentioned",
  "Customer Name": "Customer's name",
  "Issue Description": "Brief description of the problem",
  "Issue Duration": "How long the issue has been occurring",
  "Root Cause": "Suspected cause if mentioned"
}}

Transcript:
{transcript}"""
        
        try:
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that extracts information and returns valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                response_format={"type": "json_object"},
                max_tokens=500,
                temperature=0.1,
            )
            
            content = response.choices[0].message.content
            return json.loads(content)
        except Exception as e:
            st.error(f"Error extracting labels: {e}")
            return {}

class MultiAgentProcessor:
    """Multi-agent processing system with RAG integration"""
    def __init__(self, config, knowledge_service=None, cosmos_service=None):
        self.config = config
        self.knowledge_service = knowledge_service
        self.cosmos_service = cosmos_service
        self.agents = {}
        
        try:
            self.ai_client = AzureOpenAI(
                api_key=config.api_key,
                api_version=config.api_version,
                azure_endpoint=config.azure_endpoint
            )
        except Exception as e:
            st.error(f"Failed to initialize AI client: {e}")
            self.ai_client = None
        self._setup_agents()
    
    def _create_kernel(self, service_id):
        kernel = Kernel()
        kernel.add_service(
            AzureChatCompletion(
                endpoint=self.config.azure_endpoint,
                service_id=service_id,
                api_key=self.config.api_key,
                deployment_name=self.config.deployment_name,
            )
        )
        return kernel
    
    def _setup_agents(self):
        agent_configs = {
            "ManagerAgent": """You are the Manager Agent for Microsoft customer support cases. 
                             Your role is to coordinate the case resolution process:
                             1. Review incoming cases and assign them for analysis
                             2. Make decisions based on analysis results
                             3. Coordinate with other agents to ensure proper resolution
                             4. Approve final solutions before implementation
                             Keep responses concise and action-oriented.""",
            
            "AnalysisAgent": """You are the Analysis Agent for Microsoft customer support with access to historical knowledge.
                              Your role is to thoroughly analyze customer issues using both current case data and historical context:
                              1. Review case details and identify key problems
                              2. Compare with similar historical cases and their resolutions
                              3. Determine severity and impact based on past experiences
                              4. Suggest potential root causes using knowledge from previous cases
                              5. Recommend investigation approaches that have proven successful
                              Always reference relevant historical cases when making recommendations.""",
            
            "ExecutorAgent": """You are the Executor Agent for Microsoft customer support.
                              Your role is to implement solutions with detailed technical execution:
                              
                              1. **Technical Implementation**: Provide specific commands, scripts, and procedures
                              2. **Risk Assessment**: Evaluate risks and plan rollback strategies  
                              3. **Testing Protocols**: Define comprehensive testing with specific metrics
                              4. **Documentation**: Create detailed execution logs with timestamps
                              5. **Monitoring Setup**: Establish monitoring and alerting
                              6. **Validation**: Verify solutions across environments
                              
                              ALWAYS include in your response:
                              - EXECUTION_LOG: Step-by-step actions with commands
                              - VALIDATION_RESULTS: Testing outcomes with metrics
                              - ROLLBACK_PLAN: Recovery procedures if needed
                              - MONITORING_CONFIG: Setup for ongoing monitoring
                              
                              Be extremely specific about commands, configuration changes, and validation criteria.""",
            
            "NotificationAgent": """You are the Notification Agent for Microsoft customer support.
                                  Your role is to handle communications based on case outcomes:
                                  1. Create customer-friendly summaries of resolutions
                                  2. Prepare status updates for stakeholders
                                  3. Draft follow-up communications referencing resolution steps
                                  4. Ensure all parties are informed appropriately
                                  Keep communications clear, professional, and empathetic."""
        }
        
        for name, instructions in agent_configs.items():
            try:
                self.agents[name] = ChatCompletionAgent(
                    kernel=self._create_kernel(name),
                    name=name,
                    instructions=instructions,
                )
            except Exception as e:
                st.error(f"Error creating agent {name}: {e}")
    
    def _extract_issue_from_case(self, case_data):
        """Extract the core issue description for knowledge search"""
        issue_desc = case_data.get('Issue Description', '')
        root_cause = case_data.get('Root Cause', '')
        organization = case_data.get('Organization', '')
        
        search_query = f"{issue_desc} {root_cause} {organization}".strip()
        return search_query if search_query else "general support issue"
    
    def _create_resolution_summary(self, agent_responses, case_data):
        """Create a concise resolution summary for avatar"""
        summary_parts = []
        
        for response_data in agent_responses:
            agent = response_data['agent']
            response = response_data['response']
            
            if agent == "AnalysisAgent":
                summary_parts.append(f"Our analysis identified the root cause and compared it with similar historical cases.")
            elif agent == "ExecutorAgent":
                summary_parts.append(f"Our technical team implemented a comprehensive solution with proper testing and monitoring.")
            elif agent == "NotificationAgent":
                summary_parts.append(f"We've prepared detailed documentation and will provide ongoing support.")
        
        final_summary = f"""
        Your support case has been successfully resolved by our multi-agent AI system.
        
        Here's what we accomplished:
        
        • Analyzed your issue using our extensive knowledge base of similar cases
        • Implemented a comprehensive technical solution with proper testing
        • Established monitoring to prevent future occurrences
        • Documented the entire process for your reference
        
        The issue affecting {case_data.get('Organization', 'your organization')} has been fully addressed, 
        and your {case_data.get('Issue Description', 'technical issue')} is now resolved.
        
        Our team will continue monitoring to ensure stable operation.
        """
        
        return final_summary.strip()
    
    async def process_case_with_rag(self, case_data, progress_container, save_to_db=False):
        """Enhanced case processing with RAG integration"""
        if not self.agents:
            st.error("Agents not properly initialized")
            return
        
        case_id = None
        if save_to_db and self.cosmos_service:
            case_id = self.cosmos_service.save_case(case_data.copy())
            if case_id:
                st.session_state.current_case_id = case_id
        
        if self.cosmos_service and case_id:
            self.cosmos_service.log_agent_action(
                case_id, 
                "System", 
                "processing_started", 
                {"case_data": case_data, "timestamp": datetime.utcnow().isoformat()}
            )
        
        case_summary = f"""
Customer Support Case for Processing:

Customer: {case_data.get('Customer Name', 'Unknown')}
Organization: {case_data.get('Organization', 'N/A')}
Case Number: {case_data.get('Case Number', 'N/A')}
Issue: {case_data.get('Issue Description', 'No description')}
Duration: {case_data.get('Issue Duration', 'Unknown')}
Suspected Cause: {case_data.get('Root Cause', 'Not specified')}
Timestamp: {case_data.get('timestamp', 'N/A')}

Please provide your analysis and recommendations for this case.
        """
        
        agent_sequence = [
            ("ManagerAgent", "👔 Manager"),
            ("AnalysisAgent", "🔍 Analysis"), 
            ("ExecutorAgent", "⚙️ Executor"),
            ("NotificationAgent", "📧 Notification")
        ]
        
        conversation_history = case_summary
        agent_responses = []
        
        with progress_container.container():
            for agent_name, agent_emoji in agent_sequence:
                if agent_name not in self.agents:
                    st.error(f"Agent {agent_name} not found")
                    continue
                
                try:
                    st.markdown(f"""
                    <div class="agent-section">
                        <h3>{agent_emoji} {agent_name} Processing...</h3>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if agent_name == "AnalysisAgent" and self.knowledge_service:
                        agent_prompt = await self._create_rag_enhanced_prompt(case_data, conversation_history, agent_name)
                    else:
                        agent_prompt = f"""
Previous conversation:
{conversation_history}

As the {agent_name}, provide your response to this case following your role:
{self.agents[agent_name].instructions}
                        """
                    
                    response = await self._get_agent_response(agent_name, agent_prompt)
                    
                    if response:
                        st.write(response)
                        conversation_history += f"\n\n{agent_emoji} {agent_name}: {response}"
                        
                        agent_responses.append({
                            'agent': agent_name,
                            'response': response,
                            'timestamp': datetime.utcnow().isoformat()
                        })
                        
                        if self.cosmos_service and case_id:
                            self.cosmos_service.log_agent_action(
                                case_id, 
                                agent_name, 
                                "agent_response", 
                                {"response": response, "timestamp": datetime.utcnow().isoformat()}
                            )
                    else:
                        st.error(f"No response from {agent_name}")
                    
                except Exception as e:
                    st.error(f"Error with {agent_name}: {e}")
        
        resolution_summary = self._create_resolution_summary(agent_responses, case_data)
        
        if self.cosmos_service and case_id:
            self.cosmos_service.complete_case(case_id, resolution_summary)
        
        st.session_state.resolution_summary = resolution_summary
        
        # Send email notification if enabled
        if (hasattr(st.session_state, 'email_enabled') and 
            st.session_state.email_enabled and 
            hasattr(st.session_state, 'recipient_email') and 
            st.session_state.recipient_email):
            
            email_service = EmailService(self.config, show_status=False)
            if email_service.connected:
                try:
                    st.info("📧 Sending email notification...")
                    manager_name = getattr(st.session_state, 'manager_name', 'Support Manager')
                    success, message = email_service.send_case_notification(
                        case_data,
                        resolution_summary,
                        st.session_state.recipient_email,
                        manager_name
                    )
                    
                    if success:
                        st.success(f"📧 Email notification sent successfully to {st.session_state.recipient_email}")
                        st.info(f"✉️ {message}")
                    else:
                        st.error(f"❌ Failed to send email: {message}")
                except Exception as e:
                    st.error(f"❌ Email sending error: {str(e)}")
            else:
                st.warning("⚠️ Email service not available - notification not sent")
        
        st.success("✅ Case processing completed! Resolution summary is ready for avatar video.")
        
        if not save_to_db and self.cosmos_service and self.cosmos_service.connected:
            if st.button("💾 Save This Case to Database"):
                saved_case_id = self.cosmos_service.save_case(case_data.copy())
                if saved_case_id:
                    st.session_state.current_case_id = saved_case_id
                    self.cosmos_service.complete_case(saved_case_id, resolution_summary)
                    st.success(f"Case saved to database with ID: {saved_case_id}")
        
        return resolution_summary
    
    async def _create_rag_enhanced_prompt(self, case_data, conversation_history, agent_name):
        """Create RAG-enhanced prompt for Analysis Agent"""
        issue_description = self._extract_issue_from_case(case_data)
        similar_cases = self.knowledge_service.search_similar_cases(issue_description, top_k=3)
        
        base_prompt = f"""
Previous conversation:
{conversation_history}

As the {agent_name}, provide your response to this case following your role:
{self.agents[agent_name].instructions}
"""
        
        if similar_cases:
            rag_context = "\n\n**RELEVANT KNOWLEDGE FROM SEARCH:**\n"
            rag_context += "\n**Similar Historical Cases:**\n"
            
            for i, case in enumerate(similar_cases, 1):
                title = case.get('title', 'N/A')
                content = case.get('chunk', 'N/A')
                case_id = case.get('chunk_id', 'N/A')
                
                rag_context += f"{i}. **Case ID:** {case_id}\n"
                rag_context += f"   **Title:** {title}\n"
                rag_context += f"   **Content:** {content}\n\n"
            
            rag_context += """
**INSTRUCTIONS FOR USING THIS KNOWLEDGE:**
- Reference specific historical cases when they match the current issue
- Apply lessons learned from past resolutions
- Consider patterns from similar cases
- Recommend approaches that align with historical solutions
- Note any recurring themes in the retrieved cases
"""
            
            with st.expander("📚 Retrieved Knowledge (RAG Context)", expanded=False):
                st.markdown(rag_context)
            
            return base_prompt + rag_context
        else:
            st.info("No relevant historical knowledge found for this case.")
            return base_prompt
    
    async def _get_agent_response(self, agent_name, prompt):
        """Get response from a specific agent using direct API call"""
        try:
            response = self.ai_client.chat.completions.create(
                model=self.config.deployment_name,
                messages=[
                    {"role": "system", "content": self.agents[agent_name].instructions},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=800,
                temperature=0.3
            )
            return response.choices[0].message.content
        except Exception as e:
            st.error(f"API error for {agent_name}: {e}")
            return None

# Helper function for MCP operations
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
            text_content = item.get("text", "")
            lines = text_content.split('\n')
            title = lines[0] if lines else f"Result {i+1}"
            
            with st.expander(f"📄 {title[:100]}{'...' if len(title) > 100 else ''}", expanded=i==0):
                st.markdown(text_content)
        else:
            with st.expander(f"📄 Result {i+1}", expanded=i==0):
                if isinstance(item, str):
                    st.markdown(item)
                else:
                    st.json(item)

# Main application
def main():
    # Header with Microsoft styling
    st.markdown("""
    <div class="main-header">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <div class="microsoft-logo">Microsoft</div>
                <h1 style="margin: 5px 0; color: white;">AI Customer Support Agent</h1>
                <p style="margin: 0; color: #E1F5FE;">Powered by Multi-Agent AI with RAG, Avatar, Email Notifications & MCP Search</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize configuration
    config = Config()
    if not config.validate():
        st.error("Missing required environment variables. Please check your .env file configuration.")
        st.stop()
    
    # Initialize services without any status messages for clean customer experience
    cosmos_service = CosmosDBService(config, show_status=False)
    storage = InMemoryStorage(cosmos_service)
    ai_service = AIService(config)
    knowledge_service = KnowledgeService(config, show_status=False)
    avatar_service = AvatarService(config, show_status=False)  
    email_service = EmailService(config, show_status=False)
    agent_processor = MultiAgentProcessor(config, knowledge_service, cosmos_service)
    
    # Initialize MCP client in session state
    if 'mcp_client' not in st.session_state:
        st.session_state.mcp_client = None
    if 'mcp_connected' not in st.session_state:
        st.session_state.mcp_connected = False
    if 'mcp_tools' not in st.session_state:
        st.session_state.mcp_tools = []
    
    # Sidebar
    with st.sidebar:
        st.markdown("### Customer Information")
        customer_name = st.text_input("Customer Name", value="", placeholder="Enter customer name")
        support_agent = st.text_input("Support Agent", value="Microsoft Support Team")
        
        st.markdown("---")
        
        # Language selection for avatar
        st.markdown("### Avatar Language Settings")
        
        language_options = {
            "en-US": {"name": "English (US)", "voice": "en-US-AvaMultilingualNeural"},
            "es-ES": {"name": "Spanish (Spain)", "voice": "es-ES-ElviraNeural"},
            "fr-FR": {"name": "French (France)", "voice": "fr-FR-DeniseNeural"},
            "de-DE": {"name": "German (Germany)", "voice": "de-DE-KatjaNeural"},
            "it-IT": {"name": "Italian (Italy)", "voice": "it-IT-ElsaNeural"},
            "pt-BR": {"name": "Portuguese (Brazil)", "voice": "pt-BR-FranciscaNeural"},
            "zh-CN": {"name": "Chinese (Mandarin)", "voice": "zh-CN-XiaohanNeural"},
            "ja-JP": {"name": "Japanese", "voice": "ja-JP-NanamiNeural"},
            "ko-KR": {"name": "Korean", "voice": "ko-KR-SunHiNeural"},
            "hi-IN": {"name": "Hindi (India)", "voice": "hi-IN-SwaraNeural"}
        }
        
        selected_language = st.selectbox(
            "Select Avatar Language",
            options=list(language_options.keys()),
            format_func=lambda x: language_options[x]["name"],
            index=0
        )
        
        st.session_state.avatar_language = selected_language
        st.session_state.avatar_voice = language_options[selected_language]["voice"]
        
        st.markdown("---")
        
        # Email notification settings
        st.markdown("### Email Notifications")
        
        email_enabled = st.checkbox(
            "Send email notifications", 
            value=False,
            help="Automatically send email when case processing completes"
        )
        
        if email_enabled:
            recipient_email = st.text_input(
                "Recipient Email",
                placeholder="manager@company.com",
                help="Email address to receive case notifications"
            )
            
            manager_name = st.text_input(
                "Manager Name",
                value="Support Manager",
                help="Name to use in email greeting"
            )
            
            if email_service.connected:
                st.success("Email service ready")
            else:
                st.warning("Email service not configured")
        
        st.session_state.email_enabled = email_enabled
        if email_enabled:
            st.session_state.recipient_email = recipient_email
            st.session_state.manager_name = manager_name
        
        st.markdown("---")
        
        # MCP Connection Status
        st.markdown("### Microsoft Learn MCP")
        if st.session_state.mcp_connected:
            st.success("🟢 MCP Connected")
            if st.session_state.mcp_tools:
                st.info(f"📋 {len(st.session_state.mcp_tools)} tools available")
        else:
            st.error("🔴 MCP Not Connected")
            if st.button("🔌 Connect to MCP", type="secondary"):
                with st.spinner("Connecting to Microsoft Learn MCP..."):
                    client = MicrosoftLearnMCPClient()
                    success, message = run_async(client.initialize())
                    
                    if success:
                        st.session_state.mcp_client = client
                        st.session_state.mcp_connected = True
                        st.success(message)
                        
                        tools = run_async(client.list_tools())
                        st.session_state.mcp_tools = tools
                        st.rerun()
                    else:
                        st.error(message)
        
        # Quick actions
        st.markdown("---")
        st.markdown("### Quick Actions")
        
        if st.button("🧪 Test Search System"):
            if knowledge_service.search_client:
                with st.spinner("Testing search..."):
                    test_results = knowledge_service.search_similar_cases("email issues", top_k=2)
                    st.success(f"Search working - Found {len(test_results)} results")
            else:
                st.warning("Search system not available")
        
        if st.button("🎬 Check Avatar Service"):
            if avatar_service.connected:
                st.success("Avatar service ready")
            else:
                st.warning("Avatar service not available")
        
        # Clear/Reset functionality
        if st.button("🔄 Clear All Data", type="secondary"):
            # Clear all session state related to cases and processing
            keys_to_clear = [
                'cases', 'current_case_id', 'resolution_summary', 
                'last_saved_case_hash', 'current_extracted_labels',
                'mcp_search_history'
            ]
            for key in keys_to_clear:
                if key in st.session_state:
                    del st.session_state[key]
            st.success("✅ All data cleared successfully!")
            st.rerun()
        
        # Case management
        if st.session_state.cases:
            st.markdown("---")
            st.markdown(f"### Recent Cases ({len(st.session_state.cases)})")
            
            recent_cases = list(reversed(st.session_state.cases))[:3]
            for i, case in enumerate(recent_cases):
                with st.expander(f"Case {len(st.session_state.cases) - i}"):
                    st.write(f"**Customer:** {case.get('Customer Name', 'N/A')}")
                    st.write(f"**Issue:** {case.get('Issue Description', 'N/A')[:50]}...")
                    
            if len(st.session_state.cases) > 3:
                st.write(f"... and {len(st.session_state.cases) - 3} more cases")
        
        # Footer info
        st.markdown("---")
        st.markdown("**Microsoft AI Support**")
        st.caption("Powered by Azure AI Services")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Show stored cases
        if st.session_state.cases:
            with st.expander(f"📁 Case History ({len(st.session_state.cases)} cases)"):
                for i, case in enumerate(reversed(st.session_state.cases)):
                    st.write(f"**Case {len(st.session_state.cases) - i}:** {case.get('Customer Name')} - {case.get('Issue Description')}")
        
        # Transcript processing section
        st.subheader("1️⃣ Input Customer Transcript")
        
        uploaded_file = st.file_uploader("Upload transcript file", type=["txt"])
        text_input = st.text_area("Or paste transcript here:", height=200, placeholder="Paste customer support transcript here...")
        
        transcript = None
        if uploaded_file:
            transcript = StringIO(uploaded_file.getvalue().decode("utf-8")).read()
        elif text_input.strip():
            transcript = text_input
        
        # Process transcript
        if transcript:
            st.subheader("2️⃣ Extracted Information")
            
            with st.spinner("Extracting information with AI..."):
                labels = ai_service.extract_labels_from_transcript(transcript)
            
            if labels:
                # Display extracted information in cards
                info_col1, info_col2 = st.columns(2)
                
                with info_col1:
                    st.markdown(f"""
                    <div class="status-card info-card">
                        <strong>Customer:</strong> {labels.get('Customer Name', 'N/A')}<br>
                        <strong>Organization:</strong> {labels.get('Organization', 'N/A')}<br>
                        <strong>Case Number:</strong> {labels.get('Case Number', 'N/A')}
                    </div>
                    """, unsafe_allow_html=True)
                
                with info_col2:
                    st.markdown(f"""
                    <div class="status-card info-card">
                        <strong>Issue:</strong> {labels.get('Issue Description', 'N/A')}<br>
                        <strong>Duration:</strong> {labels.get('Issue Duration', 'N/A')}<br>
                        <strong>Root Cause:</strong> {labels.get('Root Cause', 'N/A')}
                    </div>
                    """, unsafe_allow_html=True)
                
                if st.button("💾 Save Case Details", type="primary"):
                    case_data = {
                        "Case Number": labels.get("Case Number", "N/A"),
                        "Organization": labels.get("Organization", "N/A"),
                        "Customer Name": labels.get("Customer Name", "N/A"),
                        "Issue Description": labels.get("Issue Description", "N/A"),
                        "Issue Duration": labels.get("Issue Duration", "N/A"),
                        "Root Cause": labels.get("Root Cause", "N/A")
                    }
                    
                    if storage.save_case(case_data):
                        st.success("✅ Case saved successfully!")
                        # Store the extracted labels to maintain the display
                        st.session_state.current_extracted_labels = labels
                        # Don't rerun to prevent losing agent communication
                        st.info("🔄 Case information is now saved and ready for processing.")
        
        # Show previously extracted labels if they exist (prevents disappearing after save)
        elif st.session_state.get('current_extracted_labels'):
            labels = st.session_state.current_extracted_labels
            
            # Display extracted information in cards
            info_col1, info_col2 = st.columns(2)
            
            with info_col1:
                st.markdown(f"""
                <div class="status-card info-card">
                    <strong>Customer:</strong> {labels.get('Customer Name', 'N/A')}<br>
                    <strong>Organization:</strong> {labels.get('Organization', 'N/A')}<br>
                    <strong>Case Number:</strong> {labels.get('Case Number', 'N/A')}
                </div>
                """, unsafe_allow_html=True)
            
            with info_col2:
                st.markdown(f"""
                <div class="status-card info-card">
                    <strong>Issue:</strong> {labels.get('Issue Description', 'N/A')}<br>
                    <strong>Duration:</strong> {labels.get('Issue Duration', 'N/A')}<br>
                    <strong>Root Cause:</strong> {labels.get('Root Cause', 'N/A')}
                </div>
                """, unsafe_allow_html=True)
            
            st.info("ℹ️ Case information already saved. You can process it with AI agents below.")
        
        # Multi-agent processing section
        st.subheader("3️⃣ Multi-Agent Case Processing")
        
        # Option 1: Process extracted case data directly
        if transcript and labels:
            if st.button("🚀 Process Current Case with AI Agents", type="primary"):
                case_data = {
                    "Case Number": labels.get("Case Number", "N/A"),
                    "Organization": labels.get("Organization", "N/A"),
                    "Customer Name": labels.get("Customer Name", "N/A"),
                    "Issue Description": labels.get("Issue Description", "N/A"),
                    "Issue Duration": labels.get("Issue Duration", "N/A"),
                    "Root Cause": labels.get("Root Cause", "N/A")
                }
                
                st.write("**Processing Case:**")
                st.json(case_data)
                
                progress_container = st.empty()
                
                with st.spinner("🤖 AI Agents processing case..."):
                    resolution_summary = asyncio.run(
                        agent_processor.process_case_with_rag(case_data, progress_container)
                    )
        
        # Option 2: Process from stored extracted labels (prevents disappearing data)
        elif st.session_state.get('current_extracted_labels'):
            stored_labels = st.session_state.current_extracted_labels
            if st.button("🚀 Process Extracted Case with AI Agents", type="primary"):
                case_data = {
                    "Case Number": stored_labels.get("Case Number", "N/A"),
                    "Organization": stored_labels.get("Organization", "N/A"),
                    "Customer Name": stored_labels.get("Customer Name", "N/A"),
                    "Issue Description": stored_labels.get("Issue Description", "N/A"),
                    "Issue Duration": stored_labels.get("Issue Duration", "N/A"),
                    "Root Cause": stored_labels.get("Root Cause", "N/A")
                }
                
                st.write("**Processing Case:**")
                st.json(case_data)
                
                progress_container = st.empty()
                
                with st.spinner("🤖 AI Agents processing case..."):
                    resolution_summary = asyncio.run(
                        agent_processor.process_case_with_rag(case_data, progress_container)
                    )
        
        # Option 3: Process saved case (if exists)
        elif st.session_state.cases:
            if st.button("🚀 Process Latest Saved Case", type="secondary"):
                latest_case = storage.fetch_latest_case()
                if latest_case:
                    st.write("**Processing Latest Saved Case:**")
                    st.json(latest_case)
                    
                    if hasattr(st.session_state, 'current_case_id') and st.session_state.current_case_id:
                        st.info(f"💾 Database Case ID: {st.session_state.current_case_id}")
                    
                    progress_container = st.empty()
                    
                    with st.spinner("🤖 AI Agents processing case..."):
                        resolution_summary = asyncio.run(
                            agent_processor.process_case_with_rag(latest_case, progress_container)
                        )
        else:
            st.info("📝 Enter a customer transcript above to extract case information, then process it directly with AI agents.")
    
    with col2:
        # Avatar section
        st.subheader("4️⃣ Generate Avatar Video")
        
        if st.session_state.get('resolution_summary'):
            st.markdown(f"""
            <div class="status-card success-card">
                ✅ Resolution summary ready for Sara avatar generation
            </div>
            """, unsafe_allow_html=True)
            
            # Show resolution summary
            with st.expander("📄 View Resolution Summary"):
                st.write(st.session_state.resolution_summary)
            
            # Manual email sending option with better status display
            if email_service.connected:
                if st.button("📧 Send Email Notification Now", type="secondary"):
                    recipient_email = st.session_state.get('recipient_email', '')
                    if not recipient_email:
                        recipient_email = st.text_input("Enter recipient email:", placeholder="manager@company.com")
                    
                    if recipient_email:
                        manager_name = st.session_state.get('manager_name', 'Support Manager')
                        latest_case = storage.fetch_latest_case()
                        
                        if latest_case:
                            try:
                                with st.spinner("Sending email notification..."):
                                    success, message = email_service.send_case_notification(
                                        latest_case,
                                        st.session_state.resolution_summary,
                                        recipient_email,
                                        manager_name
                                    )
                                    
                                    if success:
                                        st.success(f"✅ Email sent successfully to {recipient_email}")
                                        st.info(f"📧 {message}")
                                    else:
                                        st.error(f"❌ Failed to send email: {message}")
                            except Exception as e:
                                st.error(f"❌ Email sending error: {str(e)}")
                        else:
                            st.error("❌ No case data available for email")
                    else:
                        st.warning("⚠️ Please enter recipient email address")
            else:
                st.info("📧 Email service not configured - check environment variables")
            
            # Customer name for personalized avatar
            avatar_customer_name = customer_name or "Valued Customer"
            
            if st.button("🎬 Generate Avatar Video", type="primary"):
                if avatar_service.connected:
                    job_id = avatar_service._create_job_id()
                    
                    with st.spinner("🎬 Creating avatar video..."):
                        synthesis_id = avatar_service.submit_avatar_synthesis(
                            job_id, 
                            st.session_state.resolution_summary,
                            avatar_customer_name
                        )
                        
                        if synthesis_id:
                            st.info(f"📹 Avatar job submitted: {synthesis_id}")
                            
                            # Poll for completion
                            max_attempts = 60  # 5 minutes timeout
                            attempt = 0
                            
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            while attempt < max_attempts:
                                status_text.text(f"Processing... ({attempt + 1}/{max_attempts})")
                                progress_bar.progress((attempt + 1) / max_attempts)
                                
                                video_url, status_data = avatar_service.get_synthesis_status(synthesis_id)
                                
                                if video_url and status_data:
                                    if status_data.get('status') == 'Succeeded':
                                        st.success("🎉 Avatar video created successfully!")
                                        
                                        # Download and store video
                                        try:
                                            video_response = requests.get(video_url, stream=True)
                                            video_response.raise_for_status()
                                            
                                            video_name = f"support_case_{avatar_customer_name.replace(' ', '_')}_{int(time.time())}.mp4"
                                            
                                            # Upload to blob storage
                                            blob_url = avatar_service.upload_video_to_blob(
                                                video_response.content, 
                                                video_name
                                            )
                                            
                                            if blob_url:
                                                # Generate SAS token for secure access
                                                sas_token = avatar_service.generate_sas_token(video_name)
                                                
                                                if sas_token:
                                                    video_link = f"{blob_url}?{sas_token}"
                                                    
                                                    # Display video player
                                                    st.video(video_link)
                                                    
                                                    # Download button
                                                    st.download_button(
                                                        label="⬇️ Download Video",
                                                        data=video_response.content,
                                                        file_name=video_name,
                                                        mime='video/mp4'
                                                    )
                                                    
                                                    st.success("✅ Avatar video is ready!")
                                                    break
                                        except Exception as e:
                                            st.error(f"Failed to process video: {e}")
                                            break
                                    elif status_data.get('status') == 'Failed':
                                        st.error("❌ Avatar generation failed")
                                        break
                                
                                attempt += 1
                                time.sleep(5)
                            else:
                                st.error("⏰ Avatar generation timed out")
                        else:
                            st.error("❌ Failed to submit avatar job")
                else:
                    st.warning("⚠️ Avatar service not configured")
        else:
            st.markdown(f"""
            <div class="status-card warning-card">
                ⏳ Process a case first to generate resolution summary
            </div>
            """, unsafe_allow_html=True)
        
        # Video history
        if 'video_history' not in st.session_state:
            st.session_state.video_history = []
        
        if st.session_state.video_history:
            st.subheader("📹 Video History")
            for video in st.session_state.video_history:
                st.write(f"🎬 {video['name']}")

    # MCP Search section - Added after avatar video generation
    st.markdown("---")
    
    # Only show MCP section if avatar video is complete or we have a resolution summary
    if st.session_state.get('resolution_summary'):
        st.markdown("""
        <div class="mcp-section">
            <h2>5️⃣ Additional Support - Microsoft Learn Search</h2>
            <p>Need more information? Search Microsoft's official documentation for additional support.</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.session_state.mcp_connected and st.session_state.mcp_client:
            # Search interface
            st.markdown("### 🔍 Search Microsoft Documentation")
            st.markdown("Ask questions or search for additional information related to your resolved case or any Microsoft technologies.")
            
            # Create two columns for the search interface
            search_col1, search_col2 = st.columns([3, 1])
            
            with search_col1:
                # Search input
                search_query = st.text_input(
                    "Ask a question or search for information:",
                    placeholder="e.g., How to configure Azure Active Directory authentication?",
                    help="Enter your question about Microsoft technologies, services, or best practices"
                )
            
            with search_col2:
                st.markdown("<br>", unsafe_allow_html=True)  # Add spacing to align with input
                search_button = st.button("🔍 Search", type="primary", use_container_width=True)
            
            # Suggested searches based on the resolved case
            if st.session_state.cases:
                latest_case = storage.fetch_latest_case()
                if latest_case:
                    st.markdown("### 💡 Suggested Searches Based on Your Case")
                    
                    # Generate contextual suggestions
                    issue_desc = latest_case.get('Issue Description', '')
                    organization = latest_case.get('Organization', '')
                    
                    suggestions = []
                    
                    # Create smart suggestions based on issue type
                    if 'email' in issue_desc.lower():
                        suggestions.extend([
                            "Exchange Online configuration best practices",
                            "Microsoft 365 email troubleshooting guide",
                            "Outlook connectivity issues resolution"
                        ])
                    
                    if 'authentication' in issue_desc.lower() or 'login' in issue_desc.lower():
                        suggestions.extend([
                            "Azure Active Directory authentication setup",
                            "Multi-factor authentication configuration",
                            "SSO implementation with Microsoft 365"
                        ])
                    
                    if 'azure' in issue_desc.lower():
                        suggestions.extend([
                            "Azure resource management best practices",
                            "Azure monitoring and alerting setup",
                            "Azure cost optimization strategies"
                        ])
                    
                    # Default suggestions if no specific matches
                    if not suggestions:
                        suggestions = [
                            f"Microsoft solutions for {organization} enterprise needs",
                            "Microsoft 365 administration guide",
                            "Azure security best practices",
                            "Microsoft Teams deployment and management"
                        ]
                    
                    # Display suggestions in a grid
                    suggestion_cols = st.columns(2)
                    for i, suggestion in enumerate(suggestions[:4]):  # Limit to 4 suggestions
                        with suggestion_cols[i % 2]:
                            if st.button(f"💡 {suggestion}", key=f"suggestion_{i}"):
                                search_query = suggestion
                                search_button = True  # Trigger search
            
            # Perform search
            if search_button and search_query.strip():
                with st.spinner("🔍 Searching Microsoft Learn documentation..."):
                    try:
                        results = run_async(st.session_state.mcp_client.search_docs(search_query))
                        
                        st.markdown("### 📊 Search Results")
                        display_search_results(results)
                        
                        # Add search to history
                        if 'mcp_search_history' not in st.session_state:
                            st.session_state.mcp_search_history = []
                        
                        if search_query not in st.session_state.mcp_search_history:
                            st.session_state.mcp_search_history.insert(0, search_query)
                            st.session_state.mcp_search_history = st.session_state.mcp_search_history[:10]  # Keep last 10
                        
                    except Exception as e:
                        st.error(f"❌ Search failed: {str(e)}")
                        st.info("💡 Try connecting to MCP again or check your internet connection.")
            
            # Search history
            if st.session_state.get('mcp_search_history'):
                st.markdown("### 📋 Recent Searches")
                
                history_cols = st.columns(min(3, len(st.session_state.mcp_search_history)))
                
                for i, recent_query in enumerate(st.session_state.mcp_search_history[:3]):
                    with history_cols[i]:
                        if st.button(f"🔄 {recent_query[:30]}{'...' if len(recent_query) > 30 else ''}", key=f"history_{i}"):
                            with st.spinner("🔍 Searching..."):
                                results = run_async(st.session_state.mcp_client.search_docs(recent_query))
                                st.markdown(f"### Results for: {recent_query}")
                                display_search_results(results)
            
            # Help section
            with st.expander("❓ How to Use Microsoft Learn Search"):
                st.markdown("""
                **Search Tips:**
                - Ask specific questions about Microsoft technologies
                - Use clear, descriptive terms
                - Include product names (e.g., "Azure", "Microsoft 365", "Teams")
                
                **Example Queries:**
                - "How to set up Azure Functions with Python"
                - "Microsoft Teams external sharing policies"
                - "Exchange Online mailbox permissions management"
                - "Azure Active Directory conditional access setup"
                
                **What You Can Find:**
                - Official Microsoft documentation
                - Step-by-step tutorials
                - Best practices and recommendations
                - Troubleshooting guides
                - API references and code examples
                """)
        
        else:
            # MCP not connected - show connection prompt
            st.markdown("""
            <div class="status-card warning-card">
                ⚠️ Microsoft Learn MCP not connected. Connect in the sidebar to access additional documentation search.
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("### 🔗 Connect to Microsoft Learn")
            st.markdown("""
            The Microsoft Learn MCP (Model Context Protocol) provides access to:
            
            - **🔍 Semantic Search** through Microsoft's official documentation
            - **📄 Real-time Access** to the latest Microsoft Learn content
            - **🎯 Contextual Results** relevant to your support cases
            - **🏢 Official Content** from Microsoft's technical documentation
            
            Click "Connect to MCP" in the sidebar to get started.
            """)
    
    else:
        # Show placeholder when no resolution summary exists
        st.markdown("""
        <div class="mcp-section">
            <h2>5️⃣ Additional Support - Microsoft Learn Search</h2>
            <div class="status-card info-card">
                📝 Complete case processing and avatar generation first to unlock Microsoft Learn search capabilities.
            </div>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
