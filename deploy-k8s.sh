#!/bin/bash
# Excel MCP Kubernetes Deployment Script

set -e

echo "🚀 Deploying Excel MCP to Kubernetes..."

# Check if kubectl is available
if ! command -v kubectl &> /dev/null; then
    echo "❌ kubectl not found. Please install kubectl first."
    exit 1
fi

# Create namespace
echo "📁 Creating excel-mcp namespace..."
kubectl create namespace excel-mcp --dry-run=client -o yaml | kubectl apply -f -

# Apply the manifest
echo "📦 Applying Kubernetes manifest..."
kubectl apply -f excel-mcp-deployment.yaml

# Wait for deployment to be ready
echo "⏳ Waiting for deployment to be ready..."
kubectl wait --for=condition=available --timeout=300s deployment/excel-mcp

# Get service information
echo "✅ Deployment successful!"
echo ""
echo "🌐 Service URLs:"
echo "   MCP Server: http://excelgen.e-cancer.fr:31006/mcp"
echo "   File Server: http://10.2.0.150:31007/files/"
echo ""
echo "🔍 Pod status:"
kubectl get pods -l app=excel-mcp
echo ""
echo "🔍 Service status:"
kubectl get services excel-mcp-service
echo ""
echo "🧪 Test the deployment:"
echo "   curl http://excelgen.e-cancer.fr:31006/mcp"
echo ""
echo "🔗 For Open-WebUI integration, use internal DNS:"
echo "   http://excel-mcp-service.excel-mcp.svc.cluster.local:8000"
echo "   Config files: config/openwebui_*_k8s.json"