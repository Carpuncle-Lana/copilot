---
mode: 'agent'
description: 'Generate automatic SharePoint and OneDrive integration for mirroring application data using Microsoft Graph API'
---

# SharePoint/OneDrive Integration Generator

Generate a complete, portable integration solution for automatically mirroring application data files (`dashboard-status.json`, `carpuncle.log`) to SharePoint and OneDrive using Microsoft Graph API.

## Role

You are an expert Microsoft 365 integration architect with deep knowledge of Microsoft Graph API, OneDrive for Business, SharePoint Online, and cloud-agnostic application design.

## Task

Create a robust integration that:

1. **Automatically mirrors application data** to SharePoint/OneDrive
2. **Uses Microsoft Graph API** for all cloud operations
3. **Avoids local file dependencies** (no C: drive storage required)
4. **Ensures portability** and cloud-vendor independence
5. **Provides audit-proof** and transparent data exchange

## Key Requirements

### Files to Mirror

- `dashboard-status.json` - Application status dashboard data
- `carpuncle.log` - Application logging data

### Integration Principles

- **No local storage dependencies**: All operations use in-memory or temporary storage
- **Cloud-agnostic design**: Easy to swap providers if needed
- **Portable**: Works across different environments (Windows, Linux, Docker, Cloud)
- **Transparent**: Clear logging and audit trails for all operations
- **Secure**: Proper authentication and authorization using Microsoft identity platform

## Technical Approach

### Authentication Options

**Option 1: Azure AD App Registration (Recommended for production)**
- Client credentials flow (daemon/service applications)
- Certificate-based authentication
- Managed Identity (when running in Azure)

**Option 2: Delegated Permissions (Interactive scenarios)**
- OAuth 2.0 authorization code flow
- User consent for specific permissions

**Required Microsoft Graph Permissions:**
- `Files.ReadWrite.All` or `Sites.ReadWrite.All`
- For OneDrive: `Files.ReadWrite` (delegated) or `Files.ReadWrite.All` (application)
- For SharePoint: `Sites.ReadWrite.All` (application)

### Graph API Integration Patterns

**For OneDrive:**
```
PUT /me/drive/root:/path/to/file.json:/content
PUT /users/{user-id}/drive/root:/path/to/file.json:/content
PUT /drives/{drive-id}/root:/path/to/file.json:/content
```

**For SharePoint:**
```
PUT /sites/{site-id}/drive/root:/path/to/file.json:/content
PUT /sites/{site-id}/drives/{drive-id}/root:/path/to/file.json:/content
```

### Implementation Strategy

1. **Configuration Management**
   - Store credentials securely (Azure Key Vault, environment variables, managed identities)
   - Configure target SharePoint site/OneDrive location via settings
   - Define sync frequency and retry policies

2. **Data Generation & Sync Logic**
   - Generate/update `dashboard-status.json` and `carpuncle.log` in memory
   - Trigger upload on changes or at scheduled intervals
   - Handle large files with chunked uploads (>4MB)

3. **Error Handling & Resilience**
   - Implement exponential backoff for rate limiting
   - Retry logic for transient failures
   - Fallback mechanisms and offline queue

4. **Monitoring & Audit**
   - Log all sync operations with timestamps
   - Track sync status (success/failure/partial)
   - Generate audit reports for compliance

## Technology Stack Options

### Option 1: Node.js/TypeScript
- **Library**: `@microsoft/microsoft-graph-client`
- **Auth**: `@azure/identity` (DefaultAzureCredential, ClientSecretCredential)
- **Advantages**: Cross-platform, lightweight, excellent Graph SDK

### Option 2: Python
- **Library**: `msgraph-sdk-python`
- **Auth**: `azure-identity` (DefaultAzureCredential, ClientSecretCredential)
- **Advantages**: Easy scripting, great for data processing

### Option 3: C#/.NET
- **Library**: `Microsoft.Graph`
- **Auth**: `Azure.Identity` (DefaultAzureCredential, ClientSecretCredential)
- **Advantages**: Strong typing, enterprise-grade, Azure integration

### Option 4: PowerShell
- **Module**: `Microsoft.Graph` PowerShell SDK
- **Auth**: Connect-MgGraph with app registration or managed identity
- **Advantages**: Quick scripting, Windows-friendly, admin tasks

## Portable Design Patterns

### Environment-Agnostic Configuration

```yaml
# config.yml example
graph:
  tenantId: ${AZURE_TENANT_ID}
  clientId: ${AZURE_CLIENT_ID}
  clientSecret: ${AZURE_CLIENT_SECRET}  # Or use certificate
  
targets:
  sharepoint:
    siteId: ${SHAREPOINT_SITE_ID}
    driveId: ${SHAREPOINT_DRIVE_ID}
    folderPath: "/Shared Documents/ApplicationData"
    
  onedrive:
    userId: ${ONEDRIVE_USER_ID}
    folderPath: "/ApplicationSync"

sync:
  intervalSeconds: 300  # 5 minutes
  retryAttempts: 3
  chunkSizeBytes: 4194304  # 4MB
```

### Cloud-Agnostic Abstraction Layer

Create an interface/abstraction for cloud storage operations:

```typescript
interface CloudStorageProvider {
  authenticate(): Promise<void>;
  uploadFile(path: string, content: Buffer | string): Promise<void>;
  downloadFile(path: string): Promise<Buffer>;
  listFiles(path: string): Promise<FileInfo[]>;
}

class MicrosoftGraphProvider implements CloudStorageProvider {
  // Microsoft Graph implementation
}

// Future: Add GoogleDriveProvider, AWSProvider, etc.
```

## Security Best Practices

1. **Never hardcode credentials** - Use environment variables, Azure Key Vault, or managed identities
2. **Principle of least privilege** - Request only necessary Graph API permissions
3. **Encrypt data in transit** - HTTPS only (Graph API enforces this)
4. **Audit access** - Log all authentication and data access events
5. **Rotate credentials** - Regular rotation of client secrets/certificates
6. **Validate SSL certificates** - Prevent man-in-the-middle attacks

## Testing & Validation

### Unit Tests
- Mock Graph API responses
- Test authentication flows
- Validate error handling

### Integration Tests
- Test actual uploads to dev/test SharePoint site
- Verify file content integrity
- Test sync conflict resolution

### End-to-End Tests
- Full workflow from data generation to cloud sync
- Performance testing with large files
- Failure recovery scenarios

## Deployment Considerations

### Local Development
- Use dev tenant and test app registration
- Mock Graph API for offline development

### Docker Container
- Use environment variables for configuration
- No file system dependencies beyond /tmp
- Health checks for sync status

### Azure App Service / Azure Functions
- Use Managed Identity for authentication (no secrets needed)
- Application Insights for monitoring
- Azure Key Vault for sensitive configuration

### GitHub Actions / CI/CD
- Use GitHub Secrets for credentials
- Automated deployment to Azure
- Validation tests before production sync

## Example Implementation Structure

```
project/
├── src/
│   ├── auth/
│   │   ├── GraphAuthProvider.ts
│   │   └── ManagedIdentityAuth.ts
│   ├── storage/
│   │   ├── CloudStorageProvider.ts
│   │   ├── MicrosoftGraphStorage.ts
│   │   └── FileMetadata.ts
│   ├── sync/
│   │   ├── SyncEngine.ts
│   │   ├── SyncScheduler.ts
│   │   └── ConflictResolver.ts
│   ├── data/
│   │   ├── DashboardGenerator.ts
│   │   └── LogCollector.ts
│   └── index.ts
├── tests/
│   ├── unit/
│   ├── integration/
│   └── e2e/
├── config/
│   ├── config.yml
│   └── config.schema.json
├── docs/
│   ├── setup.md
│   ├── architecture.md
│   └── troubleshooting.md
├── .env.example
├── Dockerfile
└── README.md
```

## Migration Path (Avoiding Vendor Lock-in)

### Phase 1: Current State
- Microsoft Graph API integration for SharePoint/OneDrive

### Phase 2: Abstraction (Vendor Independence)
- Create CloudStorageProvider interface
- Implement Microsoft Graph provider
- All business logic uses the abstraction

### Phase 3: Multi-Cloud Ready
- Add alternative providers (Google Drive, AWS S3, etc.)
- Configuration-based provider selection
- Zero code changes to switch providers

### Phase 4: Hybrid/Multi-Cloud
- Sync to multiple clouds simultaneously
- Provider-specific optimizations
- Unified monitoring and audit

## Success Metrics

- **Reliability**: 99.9% successful sync rate
- **Performance**: Files sync within 30 seconds of generation
- **Portability**: Runs on Windows, Linux, Docker without code changes
- **Auditability**: Complete audit trail of all operations
- **Security**: Zero secrets in code, all credentials externalized

## Deliverables

When implementing this integration, provide:

1. **Complete source code** with proper error handling
2. **Configuration templates** (.env.example, config.yml.example)
3. **Setup documentation** (Azure app registration, permissions, deployment)
4. **Architecture diagram** showing data flow
5. **Test suite** (unit, integration, e2e)
6. **Monitoring dashboard** (sync status, metrics)
7. **Troubleshooting guide** (common issues and solutions)
8. **Migration guide** (for moving to alternative cloud providers)

## Advanced Features (Optional)

- **Differential sync**: Only upload changed portions of files
- **Compression**: Compress logs before upload
- **Versioning**: Maintain file history in SharePoint
- **Conflict resolution**: Handle simultaneous updates
- **Real-time sync**: WebSocket-based immediate sync
- **Bandwidth optimization**: Sync during off-peak hours
- **Geo-replication**: Multi-region SharePoint sites
- **Compliance**: GDPR, SOC2, HIPAA considerations

## Example Code Snippet (TypeScript)

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';

class SharePointSyncService {
  private graphClient: Client;

  constructor() {
    const credential = new ClientSecretCredential(
      process.env.AZURE_TENANT_ID!,
      process.env.AZURE_CLIENT_ID!,
      process.env.AZURE_CLIENT_SECRET!
    );

    this.graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          const token = await credential.getToken(
            'https://graph.microsoft.com/.default'
          );
          return token.token;
        }
      }
    });
  }

  async uploadDashboardStatus(statusData: any): Promise<void> {
    const content = JSON.stringify(statusData, null, 2);
    const filePath = '/ApplicationData/dashboard-status.json';
    
    try {
      await this.graphClient
        .api(`/sites/${process.env.SHAREPOINT_SITE_ID}/drive/root:${filePath}:/content`)
        .put(content);
      
      console.log('✅ Dashboard status synced to SharePoint');
    } catch (error) {
      console.error('❌ Failed to sync dashboard status:', error);
      throw error;
    }
  }

  async uploadLog(logContent: string): Promise<void> {
    const filePath = '/ApplicationData/carpuncle.log';
    
    try {
      await this.graphClient
        .api(`/sites/${process.env.SHAREPOINT_SITE_ID}/drive/root:${filePath}:/content`)
        .put(logContent);
      
      console.log('✅ Log file synced to SharePoint');
    } catch (error) {
      console.error('❌ Failed to sync log file:', error);
      throw error;
    }
  }
}
```

## References

- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/overview)
- [Upload files to OneDrive](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content)
- [SharePoint Sites API](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint)
- [Azure Identity SDK](https://learn.microsoft.com/en-us/dotnet/api/overview/azure/identity-readme)
- [Graph SDK for JavaScript](https://github.com/microsoftgraph/msgraph-sdk-javascript)
- [Best practices for Microsoft Graph](https://learn.microsoft.com/en-us/graph/best-practices-concept)
