"use client";

import React, { useState, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { InfoCircledIcon, CheckCircledIcon, ArrowLeftIcon } from "@radix-ui/react-icons";

interface AzureSettings {
  clientId: string;
  tenantId: string;
  redirectUri: string;
}

export default function SettingsPage() {
  const [settings, setSettings] = useState<AzureSettings>({
    clientId: "",
    tenantId: "",
    redirectUri: typeof window !== 'undefined' ? window.location.origin : "",
  });
  const [saveStatus, setSaveStatus] = useState<'idle' | 'saving' | 'saved'>('idle');

  useEffect(() => {
    // Load saved settings from localStorage
    const savedSettings = localStorage.getItem('azureSettings');
    if (savedSettings) {
      setSettings(JSON.parse(savedSettings));
    }
  }, []);

  const saveSettings = () => {
    setSaveStatus('saving');
    localStorage.setItem('azureSettings', JSON.stringify(settings));
    setSaveStatus('saved');
    setTimeout(() => setSaveStatus('idle'), 2000);
  };

  return (
    <div className="p-6 space-y-6 max-w-4xl mx-auto">
      <div className="flex items-center justify-between">
        <h1 className="text-2xl font-bold">Settings</h1>
        <Button 
          variant="outline" 
          onClick={() => window.location.href = '/'}
          className="flex items-center gap-2"
        >
          <ArrowLeftIcon className="h-4 w-4" />
          Back to Home
        </Button>
      </div>
      
      <Tabs defaultValue="azure" className="space-y-4">
        <TabsList>
          <TabsTrigger value="azure">Azure Configuration</TabsTrigger>
          <TabsTrigger value="instructions">Setup Instructions</TabsTrigger>
        </TabsList>

        <TabsContent value="azure" className="space-y-4">
          <Card>
            <CardContent className="p-6 space-y-4">
              <div className="space-y-2">
                <label className="font-semibold">Azure AD App Registration</label>
                <div className="grid gap-4">
                  <div className="space-y-2">
                    <label>Client ID (Application ID)</label>
                    <Input
                      value={settings.clientId}
                      onChange={(e) => setSettings({ ...settings, clientId: e.target.value })}
                      placeholder="Enter your Azure AD Application ID"
                    />
                  </div>
                  <div className="space-y-2">
                    <label>Tenant ID</label>
                    <Input
                      value={settings.tenantId}
                      onChange={(e) => setSettings({ ...settings, tenantId: e.target.value })}
                      placeholder="Enter your Azure AD Tenant ID"
                    />
                  </div>
                  <div className="space-y-2">
                    <label>Redirect URI</label>
                    <Input
                      value={settings.redirectUri}
                      onChange={(e) => setSettings({ ...settings, redirectUri: e.target.value })}
                      placeholder="Enter your redirect URI"
                    />
                  </div>
                </div>
                <div className="flex items-center gap-2 mt-4">
                  <Button onClick={saveSettings}>
                    {saveStatus === 'saving' ? 'Saving...' : 
                     saveStatus === 'saved' ? '✓ Saved!' : 
                     'Save Settings'}
                  </Button>
                  {saveStatus === 'saved' && (
                    <span className="text-sm text-green-600">Settings saved successfully!</span>
                  )}
                </div>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        <TabsContent value="instructions" className="space-y-4">
          <Card>
            <CardContent className="p-6 space-y-4">
              <h2 className="text-xl font-semibold">Azure AD App Registration Setup</h2>
              
              <Alert>
                <InfoCircledIcon className="h-4 w-4" />
                <AlertTitle>Required Permissions</AlertTitle>
                <AlertDescription>
                  The following Microsoft Graph API permissions are required:
                  <ul className="list-disc list-inside mt-2">
                    <li>Sites.Read.All</li>
                    <li>Sites.ReadWrite.All</li>
                    <li>User.Read</li>
                  </ul>
                </AlertDescription>
              </Alert>

              <div className="space-y-4">
                <h3 className="font-semibold">Step 1: Create App Registration</h3>
                <ol className="list-decimal list-inside space-y-2">
                  <li>Go to the Azure Portal (portal.azure.com)</li>
                  <li>Navigate to "Azure Active Directory" → "App registrations"</li>
                  <li>Click "New registration"</li>
                  <li>Enter a name for your application</li>
                  <li>Select "Single tenant" for supported account types</li>
                  <li>Add the redirect URI: {settings.redirectUri}</li>
                  <li>Click "Register"</li>
                </ol>

                <h3 className="font-semibold">Step 2: Configure Permissions</h3>
                <ol className="list-decimal list-inside space-y-2">
                  <li>In your new app registration, go to "API permissions"</li>
                  <li>Click "Add a permission"</li>
                  <li>Select "Microsoft Graph"</li>
                  <li>Choose "Delegated permissions"</li>
                  <li>Search for and add:
                    <ul className="list-disc list-inside ml-6 mt-2">
                      <li>Sites.Read.All</li>
                      <li>Sites.ReadWrite.All</li>
                      <li>User.Read</li>
                    </ul>
                  </li>
                  <li>Click "Add permissions"</li>
                  <li>Click "Grant admin consent" for your organization</li>
                </ol>

                <h3 className="font-semibold">Step 3: Get Configuration Values</h3>
                <ol className="list-decimal list-inside space-y-2">
                  <li>Copy the "Application (client) ID" to the Client ID field</li>
                  <li>Copy the "Directory (tenant) ID" to the Tenant ID field</li>
                  <li>Verify the Redirect URI matches your application's URL</li>
                </ol>

                <Alert>
                  <CheckCircledIcon className="h-4 w-4" />
                  <AlertTitle>Important Notes</AlertTitle>
                  <AlertDescription>
                    <ul className="list-disc list-inside mt-2">
                      <li>Make sure to save your settings after entering the values</li>
                      <li>The app registration must be in the same tenant as your SharePoint sites</li>
                      <li>Users will need to be granted access to the application</li>
                      <li>For cross-tenant access, you'll need to configure the app registration in both tenants</li>
                    </ul>
                  </AlertDescription>
                </Alert>
              </div>
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>
    </div>
  );
} 