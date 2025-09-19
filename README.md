

# Enterprise Trust Layer Middleware for AI Agents  

## 📌 Overview  
The **Enterprise Trust Layer Middleware for AI Agents** is a governance-first middleware designed to help enterprises deploy AI agents responsibly at scale.  

This project builds a hybrid trust layer middleware for autonomous AI agents in enterprise customer support. The middleware integrates with major CRM platforms (ServiceNow, Salesforce, Zendesk, Dynamics) via drop-in APIs. It provides real-time verification for high-risk cases (legal, compliance, billing) using lightweight ML/SLM models to block or flag untrustworthy responses before delivery. An audit dashboard ensures full traceability and empowers compliance and governance teams to oversee agentic decisions — making every AI response verifiable, auditable, and enterprise-ready.

It combines:  
- **Multi-Agent Orchestration** (Manager, Analysis, Executor, Notification)  
- **Retrieval-Augmented Generation (RAG)** with Azure AI Search + Cosmos DB  
- **Trust & Governance Controls**: Auditability, explainability, compliance, fallback mechanisms  
- **Personalization**: AI Avatars for human-like communication  
- **MCP Integration**: Real-time, verified Microsoft Learn tools & documentation  

This middleware sits at the intersection of **trust, governance, and personalization** — the three critical differentiators enterprises need for AI adoption.  

---

## ✨ Key Features  
- 🔄 **Multi-Agent Workflow** with Semantic Kernel (Manager, Analysis, Executor, Notification)  
- 📚 **RAG-enabled Knowledge Retrieval** using Azure AI Search  
- 🗂 **Case Management Persistence** with Cosmos DB (MongoDB API)  
- 🎬 **AI Avatar Video Generation** with Azure Speech & Avatar services  
- 📧 **Email Notifications** via Azure Communication Services  
- 🔍 **MCP Tools Integration** with Microsoft Learn for verified documentation  
- 🛡 **Enterprise Trust Layer**: Secure, auditable, explainable middleware  

---

## 🏗 Architecture  
flowchart TD
    A[Transcript Upload] --> B[Secure Case Ingestion]
    B --> C[Multi-Agent Processing]
    C --> D[Azure AI Search (RAG)]
    C --> E[Cosmos DB Case Store]
    C --> F[Resolution Summary]
    F --> G[AI Avatar Video]
    F --> H[Email Notifications]
    F --> I[MCP Tools Integration]

Trust Layer Pillars:

Security & Compliance
Auditability & Logging
Explainability & Transparency
Personalization (Avatar)
Verified Knowledge (MCP)

🚀 Demo Flow

Secure Case Ingestion → Up
Secure Case Ingestion → Upload transcript → Extract structured case details
Multi-Agent Processing → Manager orchestrates, Analysis uses RAG, Executor verifies, Notification drafts update
Trusted Resolution Summary → With validation logs for explainability
Personalized AI Avatar → Deliver resolution as empathetic video response
MCP Tools Integration → Search Microsoft Learn for verified knowledge

# Azure OpenAI
AZURE_OPENAI_ENDPOINT=""
AZURE_OPENAI_API_KEY=""
AZURE_OPENAI_DEPLOYMENT_NAME=""

# Azure AI Search
AZURE_SEARCH_ENDPOINT=""
AZURE_SEARCH_KEY=""
AZURE_SEARCH_INDEX=""

# Cosmos DB
AZCOSMOS_CONNSTR=""
AZCOSMOS_DATABASE_NAME=""
AZCOSMOS_CONTAINER_NAME=""

# Speech / Avatar / Blob
SPEECH_SUBSCRIPTION_KEY=""
SPEECH_ENDPOINT=""
BLOB_CONNECTION_STRING=""
BLOB_CONTAINER_NAME=""

# Email Service
AZURE_COMMUNICATION_EMAIL_CONNECTION_STRING=""
EMAIL_SENDER_ADDRESS=""

git clone https://github.com/<your-org>/<your-repo>.git
cd <your-repo>
pip install -r requirements.txt

▶️ Running the Application
streamlit run app.py

Default UI: http://localhost:8501

🤝 Contributing

Contributions are welcome! Please submit a pull request or open an issue.

📜 License

This project is licensed under the MIT License. See LICENSE
 for details.

💡 Author

Gana Chandrasekaran
Principal Technology Specialist – Azure Data & AI
