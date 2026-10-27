@description('Short name prefix for Excel engine resources')
param prefix string

@description('Azure region')
param location string = resourceGroup().location

@description('Container image for the Excel worker job')
param workerImage string = 'mcr.microsoft.com/azuredocs/containerapps-helloworld:latest'

var storageName = take(replace('${prefix}excelst', '-', ''), 24)
var sbName = take('${prefix}-excel-sb', 50)
var searchName = take(toLower('${prefix}-excel-search'), 60)
var kvName = take('${prefix}-excel-kv', 24)
var envName = take('${prefix}-excel-env', 60)
var jobName = take('${prefix}-excel-job', 60)

resource storage 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: storageName
  location: location
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    allowBlobPublicAccess: false
    minimumTlsVersion: 'TLS1_2'
  }
}

resource blobServices 'Microsoft.Storage/storageAccounts/blobServices@2023-01-01' = {
  parent: storage
  name: 'default'
}

resource container 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-01-01' = {
  parent: blobServices
  name: 'excel-workbooks'
  properties: { publicAccess: 'None' }
}

resource sb 'Microsoft.ServiceBus/namespaces@2022-10-01-preview' = {
  name: sbName
  location: location
  sku: { name: 'Basic', tier: 'Basic' }
}

resource queue 'Microsoft.ServiceBus/namespaces/queues@2022-10-01-preview' = {
  parent: sb
  name: 'excel-jobs'
  properties: {
    deadLetteringOnMessageExpiration: true
    maxDeliveryCount: 5
  }
}

resource search 'Microsoft.Search/searchServices@2023-11-01' = {
  name: searchName
  location: location
  sku: { name: 'basic' }
  properties: {
    replicaCount: 1
    partitionCount: 1
    hostingMode: 'default'
    publicNetworkAccess: 'Enabled'
  }
}

resource kv 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: kvName
  location: location
  properties: {
    tenantId: subscription().tenantId
    sku: { family: 'A', name: 'standard' }
    enableRbacAuthorization: true
    enableSoftDelete: true
  }
}

resource cae 'Microsoft.App/managedEnvironments@2023-05-01' = {
  name: envName
  location: location
  properties: {}
}

resource job 'Microsoft.App/jobs@2023-05-01' = {
  name: jobName
  location: location
  properties: {
    environmentId: cae.id
    configuration: {
      triggerType: 'Manual'
      replicaTimeout: 1800
      replicaRetryLimit: 1
      manualTriggerConfig: { parallelism: 1, replicaCompletionCount: 1 }
    }
    template: {
      containers: [
        {
          name: 'excel-engine'
          image: workerImage
          command: ['python', '-m', 'excel_engine', '--from-servicebus']
          resources: { cpu: json('1.0'), memory: '2Gi' }
          env: [
            { name: 'PULSE_EXCEL_BLOB_CONTAINER', value: 'excel-workbooks' }
            { name: 'PULSE_EXCEL_SERVICEBUS_QUEUE', value: 'excel-jobs' }
            { name: 'PULSE_EXCEL_SEARCH_ENDPOINT', value: 'https://${search.name}.search.windows.net' }
          ]
        }
      ]
    }
  }
}

output storageAccountName string = storage.name
output blobContainer string = container.name
output serviceBusNamespace string = sb.name
output serviceBusQueue string = queue.name
output searchEndpoint string = 'https://${search.name}.search.windows.net'
output keyVaultUri string = kv.properties.vaultUri
output containerAppsJob string = job.name
