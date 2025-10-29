# Kubernetes Deployment Guide

## Overview

This guide provides instructions for deploying the Excel MCP server on a Kubernetes cluster using the provided manifest.

## Prerequisites

- Kubernetes cluster (k3s recommended)
- kubectl configured to access your cluster
- Docker registry access (if using custom image)

## Quick Deployment

1. **Apply the manifest**:
   ```bash
   kubectl apply -f excel-mcp-deployment.yaml
   ```

2. **Verify deployment**:
   ```bash
   kubectl get pods
   kubectl get services
   kubectl get pvc
   ```

3. **Check logs**:
   ```bash
   kubectl logs -f deployment/excel-mcp
   ```

## External Access

The service is configured with NodePort and will be accessible at:

- **MCP Server**: `http://excelgen.e-cancer.fr:31006/mcp`
- **File Server**: `http://10.2.0.150:31007/files/`

## Open-WebUI Integration

For Open-WebUI in another namespace to connect to this service, use the internal DNS:

- **Internal DNS**: `http://excel-mcp-service.excel-mcp.svc.cluster.local:8000`

Use one of the provided Kubernetes configuration files:

- **Basic tools**: `config/openwebui_tools_k8s.json`
- **Function-based**: `config/openwebui_functions_k8s.json`
- **Enhanced tools**: `config/openwebui_tools_enhanced_k8s.json`
- **Pipe function**: `src/excel_assistant_pipe_k8s.py` (for dedicated Excel assistant model)

This file is pre-configured with the correct internal DNS and includes all available tools.

**Manual Configuration (if needed):**

```json
{
  "api_config": {
    "base_url": "http://excel-mcp-service.default.svc.cluster.local:8000",
    "endpoints": {
      "create_excel_file": {
        "url": "/mcp",
        "method": "POST",
        "headers": {
          "Content-Type": "application/json",
          "Accept": "application/json, text/event-stream"
        }
      }
    }
  }
}
```

**⚠️ Important:** Always use internal DNS (`excel-mcp-service.default.svc.cluster.local:8000`) for Open-WebUI to Excel MCP communication, never the external NodePort URLs.

## Configuration

### Environment Variables

The deployment includes a ConfigMap with the following default values:

- `HOST`: `0.0.0.0`
- `PORT`: `8000`
- `FILE_SERVER_PORT`: `8001`
- `MAX_ROWS`: `10000`
- `MAX_COLS`: `100`
- `MAX_FILENAME_LENGTH`: `255`
- `OUTPUT_DIR`: `/app/output`

### Storage

The deployment includes a 5Gi PersistentVolumeClaim for storing generated Excel files. Adjust the size as needed:

```yaml
spec:
  resources:
    requests:
      storage: 5Gi  # Change this value
```

## Troubleshooting

### Check pod status
```bash
kubectl describe pod <pod-name>
```

### Check service endpoints
```bash
kubectl get endpoints
```

### Test connectivity
```bash
# Test external access
curl http://excelgen.e-cancer.fr:31006/mcp

# Test internal DNS connectivity
./test-k8s-connection.sh
```

### View logs
```bash
kubectl logs -f deployment/excel-mcp
```

## Scaling

To scale the deployment:
```bash
kubectl scale deployment excel-mcp --replicas=3
```

## Updating

To update the image:
```bash
kubectl set image deployment/excel-mcp excel-mcp=your-new-image:tag
```

## Cleanup

To remove the deployment:
```bash
kubectl delete -f excel-mcp-deployment.yaml
```