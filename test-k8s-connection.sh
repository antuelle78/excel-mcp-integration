#!/bin/bash
# Test Kubernetes internal DNS connectivity

echo "🔍 Testing Excel MCP Kubernetes connectivity..."

# Test external access
echo "🌐 Testing external NodePort access..."
curl -s -o /dev/null -w "External MCP: %{http_code}\n" http://excelgen.e-cancer.fr:31006/mcp || echo "❌ External access failed"

# Test internal DNS (if running from within cluster)
echo "🔗 Testing internal DNS..."
kubectl run test-pod --image=curlimages/curl --rm -i --restart=Never -- curl -s http://excel-mcp-service.default.svc.cluster.local:8000/mcp || echo "❌ Internal DNS test failed"

echo "✅ Tests completed. Check pod logs for detailed results."