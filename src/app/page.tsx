// "use client" must remain at the top for React hooks to work
"use client";

import React, { useState, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Checkbox } from "@/components/ui/checkbox";
import { Card, CardContent } from "@/components/ui/card";
import { useMsal } from "@azure/msal-react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";

function AuthButtons() {
  const { instance, accounts } = useMsal();
  const [settings, setSettings] = useState<{ clientId: string; tenantId: string }>({
    clientId: "",
    tenantId: "",
  });

  useEffect(() => {
    // Load saved settings from localStorage
    const savedSettings = localStorage.getItem('azureSettings');
    if (savedSettings) {
      const parsedSettings = JSON.parse(savedSettings);
      setSettings({
        clientId: parsedSettings.clientId,
        tenantId: parsedSettings.tenantId,
      });
    }
  }, []);

  const signIn = () => {
    if (!settings.clientId || !settings.tenantId) {
      alert("Please configure your Azure AD settings first. Go to Settings page.");
      return;
    }

    instance.loginPopup({
      scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"],
      authority: `https://login.microsoftonline.com/${settings.tenantId}`,
    });
  };

  const signOut = () => {
    instance.logoutPopup();
  };

  return (
    <div className="flex items-center justify-between w-full">
      {accounts.length > 0 ? (
        <>
          <p>Signed in as: {accounts[0].username}</p>
          <div className="flex gap-2">
            <Button onClick={signOut}>Sign out</Button>
            <Button variant="outline" onClick={() => window.location.href = '/settings'}>
              ⚙️ Settings
            </Button>
          </div>
        </>
      ) : (
        <div className="flex gap-2">
          <Button onClick={signIn}>Sign in with Microsoft</Button>
          <Button variant="outline" onClick={() => window.location.href = '/settings'}>
            ⚙️ Settings
          </Button>
        </div>
      )}
    </div>
  );
}

interface SiteItem {
  name: string;
  type: "List" | "Document Library" | "Page";
  selected: boolean;
}

interface ListColumn {
  Title: string;
  FieldTypeKind: number;
  Required: boolean;
  Hidden: boolean;
  ReadOnlyField: boolean;
  __metadata?: { type: string };
}

interface ListField {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  Required: boolean;
  Hidden: boolean;
  ReadOnlyField: boolean;
  FieldTypeKind: number;
  Choices?: string[];
  DefaultValue?: string;
  LookupList?: string;
  LookupField?: string;
  __metadata?: {
    type: string;
  };
}

interface ListView {
  Title: string;
  ViewFields: {
    results: string[];
  };
  ViewQuery: string;
  RowLimit: number;
  Paged: boolean;
  DefaultView: boolean;
  __metadata: {
    type: string;
    uri?: string;
    id?: string;
  };
  Id?: string;
}

interface ListItemBasic {
  Id: number;
  Title: string;
  __metadata?: {
    type: string;
    uri: string;
  };
}

interface ListItemFull {
  Id: number;
  Title: string;
  [key: string]: any;
}

interface SharePointResponse<T> {
  d: {
    results: T[];
  };
}

interface GraphListItem {
  name: string;
  list?: {
    template: string;
  };
}

interface GraphPageItem {
  name: string;
}

// Helper functions for list operations
async function deleteListIfExists(hostname: string, path: string, listName: string, accessToken: string) {
  try {
    const checkListRes = await fetch(
      `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (checkListRes.ok) {
      const deleteListRes = await fetch(
        `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')`,
        {
          method: "DELETE",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json;odata=verbose",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
          },
        }
      );

      if (deleteListRes.ok) {
        return true;
      }
    }
    return false;
  } catch (error) {
    return false;
  }
}

async function createList(hostname: string, path: string, listName: string, baseTemplate: number, accessToken: string) {
  const createRes = await fetch(
    `https://${hostname}${path}/_api/web/lists`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
      },
      body: JSON.stringify({
        __metadata: { type: "SP.List" },
        Title: listName,
        BaseTemplate: baseTemplate,
      }),
    }
  );

  if (!createRes.ok) {
    throw new Error(`Failed to create list: ${await createRes.text()}`);
  }

  return await createRes.json();
}

async function createListField(hostname: string, path: string, listName: string, field: ListField, accessToken: string) {
  try {
    // Use AddFieldAsXml for Choice fields
    if (field.TypeAsString === 'Choice' && field.Choices) {
      // Support both array and object-with-results for Choices
      const choicesArray = Array.isArray(field.Choices)
        ? field.Choices
        : (typeof field.Choices === 'object' && field.Choices && 'results' in field.Choices && Array.isArray((field.Choices as any).results)
            ? (field.Choices as any).results
            : []);
      const choicesXml = choicesArray.map((choice: string) => `<CHOICE>${choice}</CHOICE>`).join('');
      const fieldXml = `
        <Field Type="Choice" Name="${field.InternalName}" DisplayName="${field.Title}" Format="Dropdown">
          <CHOICES>${choicesXml}</CHOICES>
        </Field>
      `;
      const xmlPayload = {
        parameters: {
          __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
          SchemaXml: fieldXml,
          Options: 0
        }
      };
      const createFieldRes = await fetch(
        `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/fields/CreateFieldAsXml`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
          },
          body: JSON.stringify(xmlPayload),
        }
      );
      if (!createFieldRes.ok) {
        throw new Error(`Failed to create field: ${await createFieldRes.text()}`);
      }
      return;
    }
    // Fallback to REST for other field types
    let fieldPayload: any = {
      __metadata: { type: "SP.Field" },
      Title: field.Title,
      FieldTypeKind: field.FieldTypeKind,
      Required: field.Required,
    };
    switch (field.TypeAsString) {
      case 'User':
        fieldPayload = {
          ...fieldPayload,
          __metadata: { type: "SP.FieldUser" },
          SelectionMode: 1 // 1 for single user, 2 for multiple users
        };
        break;
      case 'DateTime':
        fieldPayload = {
          ...fieldPayload,
          __metadata: { type: "SP.FieldDateTime" },
          DisplayFormat: 0, // 0 for date only, 1 for date and time
          DateTimeCalendarType: 0, // 0 for Gregorian
          FriendlyDisplayFormat: 0 // 0 for default
        };
        break;
      case 'Lookup':
        fieldPayload = {
          ...fieldPayload,
          __metadata: { type: "SP.FieldLookup" },
          LookupList: field.LookupList,
          LookupField: field.LookupField
        };
        break;
    }
    if (field.DefaultValue) {
      fieldPayload.DefaultValue = field.DefaultValue;
    }
    const createFieldRes = await fetch(
      `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/fields`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify(fieldPayload),
      }
    );
    if (!createFieldRes.ok) {
      throw new Error(`Failed to create field: ${await createFieldRes.text()}`);
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
    throw new Error(`Failed to create field: ${errorMessage}`);
  }
}

async function createListView(
  hostname: string,
  path: string,
  listName: string,
  view: ListView,
  accessToken: string,
  setLogs: (updater: (prev: string[]) => string[]) => void,
  customFormatter?: string,
  destFields?: ListField[]
) {
  try {
    // 1. Create the view with basic properties
    const viewPayload = {
      __metadata: { type: "SP.View" },
      Title: view.Title,
      PersonalView: false,
      ViewQuery: view.ViewQuery || "",
      RowLimit: view.RowLimit || 30,
      Paged: view.Paged !== false,
      DefaultView: view.DefaultView || false
    };

    setLogs((prev) => [...prev, `DEBUG: Creating view: ${view.Title}`]);

    const createViewRes = await fetch(
      `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/views`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify(viewPayload),
      }
    );

    if (!createViewRes.ok) {
      const errorText = await createViewRes.text();
      setLogs((prev) => [...prev, `DEBUG: View creation error: ${errorText}`]);
      throw new Error(`Failed to create view: ${errorText}`);
    }

    const viewData = await createViewRes.json();
    const viewId = viewData.d.Id;

    // 2. Remove all fields from the view
    const removeFieldsRes = await fetch(
      `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/views('${viewId}')/ViewFields/RemoveAllViewFields`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose"
        }
      }
    );
    if (!removeFieldsRes.ok) {
      removeFieldsRes.text().then(errorText => {
        setLogs((prev) => [...prev, `⚠️ Failed to remove all fields from view: ${errorText}`]);
      }).catch(() => {
        setLogs((prev) => [...prev, `⚠️ Failed to remove all fields from view: (error reading response)`]);
      });
    }

    // 3. Map source field names to destination internal names
    const sourceFields = view.ViewFields?.results || [];
    let fieldNameMap: Record<string, string> = {};
    if (destFields) {
      setLogs((prev) => [
        ...prev,
        `DEBUG: All destination field InternalNames: ${destFields.map(f => f.InternalName).join(', ')}`
      ]);
      setLogs((prev) => [
        ...prev,
        ...destFields.map(f => `DEBUG: Field object: ${JSON.stringify(f)}`)
      ]);
      for (const f of destFields) {
        // Add the actual InternalName as a key
        fieldNameMap[f.InternalName] = f.InternalName;
        // Add encoded internal name variants as keys
        fieldNameMap[(f.InternalName || '').replace(/_x0020_/g, '')] = f.InternalName;
        fieldNameMap[(f.InternalName || '').replace(/_x0020_/g, '').toLowerCase()] = f.InternalName;
        const staticName = (f as any).StaticName || '';
        const names = [
          f.InternalName,
          staticName,
          f.Title,
          (f.InternalName || '').toLowerCase(),
          (staticName || '').toLowerCase(),
          (f.Title || '').toLowerCase(),
          (f.Title || '').replace(/\s/g, ''),
          (f.Title || '').replace(/\s/g, '').toLowerCase(),
          (f.InternalName || '').replace(/_x0020_/g, ''),
          (f.InternalName || '').replace(/_x0020_/g, '').toLowerCase(),
        ];
        for (const name of names) {
          if (name) fieldNameMap[name] = f.InternalName;
        }
      }
      setLogs((prev) => [
        ...prev,
        `DEBUG: Destination field name map: ${JSON.stringify(fieldNameMap)}`
      ]);
    }

    // Build a map of all possible internal names for each title
    let titleToInternalNames: Record<string, string[]> = {};
    if (destFields) {
      for (const f of destFields) {
        const normalizedTitle = (f.Title || '').replace(/\s/g, '').toLowerCase();
        if (!titleToInternalNames[normalizedTitle]) titleToInternalNames[normalizedTitle] = [];
        titleToInternalNames[normalizedTitle].push(f.InternalName);
      }
      setLogs((prev) => [
        ...prev,
        `DEBUG: Title to InternalNames map: ${JSON.stringify(titleToInternalNames)}`
      ]);
    }

    // Helper to normalize field names
    const normalized = (name: string) => name.replace(/\s/g, '').replace(/_x0020_/g, '').toLowerCase();

    // 4. Add only the mapped fields, in order
    for (const field of sourceFields) {
      let destFieldName: string | null = null;
      const normalizedField = field.replace(/\s/g, '').toLowerCase();

      // First try: Use hardcoded mappings for known fields with spaces
      const hardcodedMappings: Record<string, string> = {
        'StartDate': 'Start_x0020_Date',
        'Start Date': 'Start_x0020_Date',
        'EndDate': 'End_x0020_Date',
        'End Date': 'End_x0020_Date',
        'Assignedto': 'Assigned_x0020_to',
        'Assigned To': 'Assigned_x0020_to',
        'Assigned to': 'Assigned_x0020_to'
      };

      if (hardcodedMappings[field]) {
        destFieldName = hardcodedMappings[field];
        setLogs((prev) => [
          ...prev,
          `DEBUG: Using hardcoded mapping for field '${field}' -> '${destFieldName}'`
        ]);
      }

      // Second try: Prefer encoded internal name if available
      if (!destFieldName && titleToInternalNames[normalizedField]) {
        const allNames = titleToInternalNames[normalizedField] || [];
        // Sort so encoded names (_x0020_) come first
        allNames.sort((a, b) => (b.includes('_x0020_') ? 1 : -1) - (a.includes('_x0020_') ? 1 : -1));
        if (allNames.length > 0) {
          destFieldName = allNames[0];
          setLogs((prev) => [
            ...prev,
            `DEBUG: For view field '${field}', using sorted internal name '${destFieldName}' from titleToInternalNames map (all: ${JSON.stringify(allNames)})`
          ]);
        }
      }

      // Third try: Previous mapping logic
      if (!destFieldName) {
        destFieldName =
          fieldNameMap[field] ||
          fieldNameMap[field.toLowerCase()] ||
          fieldNameMap[normalized(field)] ||
          null;
      }

      // Fourth try: match by Title (spaces removed, lowercased)
      if (!destFieldName && destFields) {
        const matchByTitle = destFields.find(
          f =>
            f.Title &&
            f.Title.replace(/\s/g, '').toLowerCase() === field.replace(/\s/g, '').toLowerCase()
        );
        if (matchByTitle) {
          destFieldName = matchByTitle.InternalName;
          setLogs((prev) => [
            ...prev,
            `DEBUG: Fallback matched view field '${field}' to destination Title '${matchByTitle.Title}' with internal name '${destFieldName}'`
          ]);
        }
      }

      // Fifth try: match by encoded InternalName (remove _x0020_ and compare)
      if (!destFieldName && destFields) {
        const matchByEncoded = destFields.find(
          f =>
            f.InternalName &&
            f.InternalName.replace(/_x0020_/g, '').toLowerCase() === field.replace(/\s/g, '').toLowerCase()
        );
        if (matchByEncoded) {
          destFieldName = matchByEncoded.InternalName;
          setLogs((prev) => [
            ...prev,
            `DEBUG: Fallback matched view field '${field}' to encoded InternalName '${matchByEncoded.InternalName}'`
          ]);
        }
      }

      // Final fallback: use the original field name
      if (!destFieldName) {
        destFieldName = field;
        setLogs((prev) => [
          ...prev,
          `DEBUG: Using original field name as fallback: '${field}'`
        ]);
      }

      setLogs((prev) => [...prev, `DEBUG: Final destFieldName for view field '${field}': '${destFieldName}'`]);

      try {
        const addFieldRes = await fetch(
          `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/views('${viewId}')/ViewFields/AddViewField('${destFieldName}')`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose"
            }
          }
        );
        if (!addFieldRes.ok) {
          const errorText = await addFieldRes.text();
          setLogs((prev: string[]) => [...prev, `⚠️ Failed to add field ${destFieldName} to view: ${errorText}`]);
        } else {
          setLogs((prev: string[]) => [...prev, `✅ Added field ${destFieldName} to view`]);
        }
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev: string[]) => [...prev, `⚠️ Failed to add field ${destFieldName} to view: ${errorMessage}`]);
      }
    }

    // PATCH CustomFormatter if provided
    if (customFormatter) {
      const patchRes = await fetch(
        `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/views('${viewId}')`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-HTTP-Method": "MERGE"
          },
          body: JSON.stringify({ __metadata: { type: "SP.View" }, CustomFormatter: customFormatter })
        }
      );
      const patchText = await patchRes.text();
      setLogs((prev: string[]) => [...prev, `DEBUG: CustomFormatter PATCH response for view ${view.Title}: ${patchText}`]);
      if (!patchRes.ok) {
        setLogs((prev: string[]) => [...prev, `⚠️ Failed to patch CustomFormatter: ${patchText}`]);
      } else {
        setLogs((prev: string[]) => [...prev, `✅ Patched CustomFormatter for view: ${view.Title}`]);
      }
    }

    setLogs((prev: string[]) => [...prev, `✅ Created view: ${view.Title} with ${sourceFields.length} fields`]);
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
    throw new Error(`Failed to create view: ${errorMessage}`);
  }
}

async function createListItem(hostname: string, path: string, listName: string, item: ListItemBasic, accessToken: string) {
  try {
    const cleanItem = Object.entries(item)
      .filter(([key, value]) => !key.startsWith('_') && !key.startsWith('__') && value !== undefined && value !== null)
      .reduce((acc, [key, value]) => {
        acc[key] = value;
        return acc;
      }, {} as Record<string, any>);

    const itemPayload = {
      __metadata: { type: "SP.ListItem" },
      ...cleanItem
    };

    const createItemRes = await fetch(
      `https://${hostname}${path}/_api/web/lists/GetByTitle('${listName}')/items`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify(itemPayload),
      }
    );

    if (!createItemRes.ok) {
      throw new Error(`Failed to create item: ${await createItemRes.text()}`);
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
    throw new Error(`Failed to create item: ${errorMessage}`);
  }
}

// Main cloning function
async function cloneList(
  sourceHostname: string,
  sourcePath: string,
  targetHostname: string,
  targetPath: string,
  listName: string,
  accessToken: string,
  setLogs: (updater: (prev: string[]) => string[]) => void
) {
  try {
    // Delete existing list if it exists
    const wasDeleted = await deleteListIfExists(targetHostname, targetPath, listName, accessToken);
    if (wasDeleted) {
      setLogs((prev) => [...prev, `Deleted existing list: ${listName}`]);
    }

    // Get list details from source
    const encodedListTitle = encodeURIComponent(listName);
    const restListDetailsRes = await fetch(
      `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (!restListDetailsRes.ok) {
      throw new Error(`Failed to get list details: ${await restListDetailsRes.text()}`);
    }

    const restListDetails = await restListDetailsRes.json();
    if (!restListDetails?.d?.BaseTemplate) {
      throw new Error('BaseTemplate not found in list details');
    }

    // Create the list
    await createList(targetHostname, targetPath, listName, restListDetails.d.BaseTemplate, accessToken);
    setLogs((prev) => [...prev, `✅ Created list: ${listName}`]);

    // Get and create fields
    const fieldsRes = await fetch(
      `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/fields`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (!fieldsRes.ok) {
      throw new Error(`Failed to fetch fields: ${await fieldsRes.text()}`);
    }

    const fieldsData = await fieldsRes.json() as SharePointResponse<ListField>;
    const builtInFields = ['ContentType', 'Attachments', 'Title', 'ID', 'Created', 'Modified', 'Author', 'Editor'];
    const fields = fieldsData.d.results.filter(field => 
      !field.Hidden && 
      !field.ReadOnlyField && 
      !builtInFields.includes(field.InternalName)
    );

    // Create each field
    for (const field of fields) {
      try {
        await createListField(targetHostname, targetPath, listName, field, accessToken);
        setLogs((prev) => [...prev, `✅ Created field: ${field.Title}`]);
        setLogs((prev) => [...prev, `DEBUG: Dest field: ${field.Title} (${field.InternalName})`]);
        // Fetch CustomFormatter from source (explicitly select the property)
        const fieldDetailsRes = await fetch(
          `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/fields/GetByInternalNameOrTitle('${field.InternalName}')?$select=CustomFormatter`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json;odata=verbose",
            },
          }
        );
        if (fieldDetailsRes.ok) {
          const fieldDetails = await fieldDetailsRes.json();
          const customFormatter = fieldDetails.d.CustomFormatter;
          setLogs((prev) => [...prev, `DEBUG: Source CustomFormatter for ${field.InternalName}: ${customFormatter}`]);
          if (customFormatter) {
            // PATCH CustomFormatter to the destination field
            const patchRes = await fetch(
              `https://${targetHostname}${targetPath}/_api/web/lists/GetByTitle('${listName}')/fields/GetByInternalNameOrTitle('${field.InternalName}')`,
              {
                method: "POST",
                headers: {
                  Authorization: `Bearer ${accessToken}`,
                  Accept: "application/json;odata=verbose",
                  "Content-Type": "application/json;odata=verbose",
                  "X-HTTP-Method": "MERGE"
                },
                body: JSON.stringify({ __metadata: { type: "SP.Field" }, CustomFormatter: customFormatter })
              }
            );
            const patchText = await patchRes.text();
            setLogs((prev) => [...prev, `DEBUG: CustomFormatter PATCH response for ${field.InternalName}: ${patchText}`]);
            if (!patchRes.ok) {
              setLogs((prev) => [...prev, `⚠️ Failed to patch CustomFormatter for field ${field.Title}: ${patchText}`]);
            } else {
              setLogs((prev) => [...prev, `✅ Patched CustomFormatter for field: ${field.Title}`]);
            }
          }
        }
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev) => [...prev, `⚠️ Failed to create field ${field.Title}: ${errorMessage}`]);
      }
    }

    // Add a delay to allow SharePoint to finish provisioning fields
    setLogs((prev) => [...prev, 'DEBUG: Waiting 3 seconds for SharePoint field propagation...']);
    await new Promise(res => setTimeout(res, 3000));

    // Get and create views
    const viewsRes = await fetch(
      `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/views`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (!viewsRes.ok) {
      throw new Error(`Failed to fetch views: ${await viewsRes.text()}`);
    }

    const viewsData = await viewsRes.json() as SharePointResponse<ListView>;
    const views = viewsData.d.results;

    // Fetch and log the list fields again before creating each view
    const destFieldsRes = await fetch(
      `https://${targetHostname}${targetPath}/_api/web/lists/GetByTitle('${listName}')/fields`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );
    if (destFieldsRes.ok) {
      const destFieldsData = await destFieldsRes.json();
      setLogs((prev) => [
        ...prev,
        'DEBUG: Fields in destination list just before view creation:',
        ...destFieldsData.d.results.map((f: any) => `DEBUG: Field object before view: ${JSON.stringify(f)}`)
      ]);
    }

    // Create each view
    for (const view of views) {
      try {
        // Fetch view details including fields using $expand
        const viewDetailsRes = await fetch(
          `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/views('${view.Id}')?$select=Title,ViewQuery,RowLimit,Paged,DefaultView,CustomFormatter&$expand=ViewFields`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json;odata=verbose",
            },
          }
        );
        
        if (!viewDetailsRes.ok) {
          throw new Error(`Failed to fetch view details: ${await viewDetailsRes.text()}`);
        }

        const viewDetails = await viewDetailsRes.json();
        setLogs((prev) => [...prev, `DEBUG: Full view details: ${JSON.stringify(viewDetails.d)}`]);
        
        // Parse the SchemaXml to get field names
        const schemaXml = viewDetails.d.ViewFields?.SchemaXml || '';
        const fieldMatches = schemaXml.match(/Name="([^"]+)"/g) || [];
        const viewFields = fieldMatches.map((match: string) => match.match(/Name="([^"]+)"/)?.[1]).filter(Boolean);
        
        setLogs((prev) => [...prev, `DEBUG: Parsed view fields from SchemaXml: ${JSON.stringify(viewFields)}`]);
        
        // Fetch CustomFormatter for this view
        const customFormatter = viewDetails.d.CustomFormatter;
        setLogs((prev) => [...prev, `DEBUG: Source CustomFormatter for view ${view.Title}: ${customFormatter}`]);

        // Create the view with the fetched details
        await createListView(
          targetHostname, 
          targetPath, 
          listName, 
          {
            ...view,
            ViewFields: {
              results: viewFields
            }
          }, 
          accessToken, 
          setLogs, 
          customFormatter,
          fields
        );
        setLogs((prev) => [...prev, `✅ Created view: ${view.Title}`]);
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev) => [...prev, `⚠️ Failed to create view ${view.Title}: ${errorMessage}`]);
      }
    }

    // Build the select and expand query parameters
    const selectFields = ['Title', ...fields.map(f => f.InternalName)];
    let expandFields = '';
    
    // Always include Assignedto/Id, Assignedto/Title, Assignedto/EMail if Assignedto field exists
    const assignedToField = fields.find(f => f.InternalName === 'Assignedto');
    if (assignedToField) {
      selectFields.push('Assignedto/Id', 'Assignedto/Title', 'Assignedto/EMail');
      expandFields = 'Assignedto';
    }
    const selectParam = `$select=Id,${selectFields.join(',')}`;
    const expandParam = expandFields ? `&$expand=${expandFields}` : '';

    // Use these in the fetch URL for items
    const itemsRes = await fetch(
      `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/items?${selectParam}${expandParam}`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json;odata=verbose",
        },
      }
    );

    if (!itemsRes.ok) {
      const errorText = await itemsRes.text();
      setLogs((prev) => [...prev, `⚠️ Failed to fetch items: ${errorText}`]);
      throw new Error(`Failed to fetch items: ${errorText}`);
    }

    const itemsData = await itemsRes.json() as SharePointResponse<ListItemBasic>;
    if (!itemsData.d?.results) {
      setLogs((prev) => [...prev, "No items found in source list"]);
      return;
    }

    const items = itemsData.d.results;
    setLogs((prev) => [...prev, `Found ${items.length} items to copy`]);

    // Create each item
    for (const item of items) {
      try {
        if (!item.Id) {
          setLogs((prev) => [...prev, `⚠️ Skipping item '${item.Title}' - no ID found`]);
          continue;
        }
        // Get the full item data including all custom fields
        const fullItemRes = await fetch(
          `https://${sourceHostname}${sourcePath}/_api/web/lists/GetByTitle('${encodedListTitle}')/items(${item.Id})?${selectParam}${expandParam}`,
          {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json;odata=verbose",
            },
          }
        );
        if (!fullItemRes.ok) {
          const errorText = await fullItemRes.text();
          setLogs((prev) => [...prev, `⚠️ Failed to fetch full item data for item ${item.Id}: ${errorText}`]);
          continue;
        }
        const fullItem = await fullItemRes.json() as { d: ListItemFull };
        setLogs((prev) => [...prev, `DEBUG: Source item data: ${JSON.stringify(fullItem.d)}`]);
        const fieldNameMap = {
          StartDate: "Start_x0020_Date",
          EndDate: "End_x0020_Date",
          AssignedtoId: "Assigned_x0020_toId"
        };
        const cleanItem = await Object.entries(fullItem.d)
          .filter(([key, value]) => {
            return !key.startsWith('_') && !key.startsWith('__') && value !== undefined && value !== null;
          })
          .reduce(async (accPromise, [key, value]) => {
            const acc = await accPromise;
            if (key === 'Assignedto' && value && typeof value === 'object') {
              // Try to get login name, email, or title
              let loginName = value.LoginName;
              if (!loginName && value.EMail) loginName = value.EMail;
              if (!loginName && value.Title) loginName = value.Title;
              // If still no loginName, try to fetch from getuserbyid
              if (!loginName && value.Id) {
                const userInfoRes = await fetch(
                  `https://${sourceHostname}${sourcePath}/_api/web/getuserbyid(${value.Id})`,
                  {
                    headers: {
                      Authorization: `Bearer ${accessToken}`,
                      Accept: "application/json;odata=verbose",
                    },
                  }
                );
                if (userInfoRes.ok) {
                  const userInfo = await userInfoRes.json();
                  loginName = userInfo.d.LoginName || userInfo.d.Email || userInfo.d.Title;
                }
              }
              if (loginName) {
                // Ensure user in destination and get ID
                const ensureUserRes = await fetch(
                  `https://${targetHostname}${targetPath}/_api/web/ensureuser`,
                  {
                    method: 'POST',
                    headers: {
                      Authorization: `Bearer ${accessToken}`,
                      Accept: 'application/json;odata=verbose',
                      'Content-Type': 'application/json;odata=verbose'
                    },
                    body: JSON.stringify({ logonName: loginName })
                  }
                );
                if (ensureUserRes.ok) {
                  const ensureUserData = await ensureUserRes.json();
                  setLogs((prev) => [...prev, `DEBUG: ensureuser response: ${JSON.stringify(ensureUserData.d)}`]);
                  acc[(fieldNameMap as Record<string, string>)['AssignedtoId']] = ensureUserData.d.Id;
                }
              }
            } else if (key !== 'Assignedto') {
              const destKey = (fieldNameMap as Record<string, string>)[key] || key;
              acc[destKey] = value;
            }
            return acc;
          }, Promise.resolve({} as Record<string, any>));
        const itemPayload = {
          __metadata: { type: "SP.ListItem" },
          ...cleanItem
        };
        // Debug log the payload
        setLogs((prev) => [...prev, `DEBUG: Creating item payload: ${JSON.stringify(itemPayload)}`]);
        const createItemRes = await fetch(
          `https://${targetHostname}${targetPath}/_api/web/lists/GetByTitle('${listName}')/items`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${accessToken}`,
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
            },
            body: JSON.stringify(itemPayload),
          }
        );
        const createItemText = await createItemRes.text();
        setLogs((prev) => [...prev, `DEBUG: create item response: ${createItemText}`]);
        if (createItemRes.ok) {
          setLogs((prev) => [...prev, `✅ Created item: ${item.Title || 'Untitled'}`]);
        } else {
          setLogs((prev) => [...prev, `⚠️ Failed to create item: ${createItemText}`]);
        }
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev) => [...prev, `❌ Error creating item: ${errorMessage}`]);
      }
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
    throw new Error(`Error cloning list '${listName}': ${errorMessage}`);
  }
}

export default function SharePointCloner() {
  const { instance, accounts } = useMsal();

  const [sourceUrl, setSourceUrl] = useState("");
  const [targetUrl, setTargetUrl] = useState("");
  const [cloneLists, setCloneLists] = useState(true);
  const [cloneLibraries, setCloneLibraries] = useState(true);
  const [clonePages, setClonePages] = useState(true);
  const [cloneSite, setCloneSite] = useState(false);
  const [logs, setLogs] = useState<string[]>(["Waiting for user..."]);
  const [items, setItems] = useState<SiteItem[]>([]);

  const toggleItem = (index: number) => {
    const updated = [...items];
    updated[index].selected = !updated[index].selected;
    setItems(updated);
  };

  const loadItemsFromSite = async () => {
    if (accounts.length === 0) {
      setLogs((prev) => [...prev, "User not signed in."]);
      return;
    }

    try {
      const response = await instance.acquireTokenSilent({
        scopes: ["Sites.Read.All"],
        account: accounts[0],
      });

      const accessToken = response.accessToken;
      const siteHostname = new URL(sourceUrl).hostname;
      const sitePath = new URL(sourceUrl).pathname;

      const siteResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteHostname}:${sitePath}`,
        {
          headers: { Authorization: `Bearer ${accessToken}` },
        }
      );

      const siteData = await siteResp.json();
      const siteId = siteData.id;

      const [listsResp, pagesResp] = await Promise.all([
        fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists`, {
          headers: { Authorization: `Bearer ${accessToken}` },
        }),
        fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/pages`, {
          headers: { Authorization: `Bearer ${accessToken}` },
        }),
      ]);

      const listData = await listsResp.json();
      const pageData = await pagesResp.json();

      const listItems = listData.value.map((list: GraphListItem) => ({
        name: list.name,
        type: list.list?.template === "documentLibrary" ? "Document Library" : "List",
        selected: true,
      }));

      const pageItems = pageData.value.map((page: GraphPageItem) => ({
        name: page.name,
        type: "Page",
        selected: true,
      }));

      const combinedItems = [...listItems, ...pageItems];

      setItems(combinedItems);
      setLogs((prev) => [...prev, `Loaded ${combinedItems.length} items from source site.`]);
    } catch (error: unknown) {
      if (error instanceof InteractionRequiredAuthError) {
        instance.acquireTokenPopup({
          scopes: ["Sites.Read.All"],
        });
      } else {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
        setLogs((prev) => [...prev, `Error loading site items: ${errorMessage}`]);
      }
    }
  };

  const startCloning = async () => {
    setLogs((prev) => [...prev, "Starting clone..."]);

    if (accounts.length === 0) {
      setLogs((prev) => [...prev, "Cannot clone: user not signed in."]);
      return;
    }

    try {
      const sourceHostname = new URL(sourceUrl).hostname;
      const sourcePath = new URL(sourceUrl).pathname;
      const targetHostname = new URL(targetUrl).hostname;
      const targetPath = new URL(targetUrl).pathname;
      
      const response = await instance.acquireTokenSilent({
        scopes: [`https://${sourceHostname}/.default`],
        account: accounts[0],
      });

      const accessToken = response.accessToken;

      for (const item of items) {
        if (!item.selected) continue;

        if (item.type === "List" && cloneLists) {
          setLogs((prev) => [...prev, `Cloning list: ${item.name}`]);
          try {
            await cloneList(sourceHostname, sourcePath, targetHostname, targetPath, item.name, accessToken, setLogs);
          } catch (error: unknown) {
            const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
            setLogs((prev) => [...prev, `❌ ${errorMessage}`]);
          }
        } else if (item.type === "Document Library" && cloneLibraries) {
          setLogs((prev) => [...prev, `Cloning library: ${item.name}`]);
          // Placeholder: implement document library clone logic
        } else if (item.type === "Page" && clonePages) {
          setLogs((prev) => [...prev, `Cloning page: ${item.name}`]);
          // Placeholder: implement page cloning
        }
      }

      setLogs((prev) => [...prev, "✅ Cloning completed."]);
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      setLogs((prev) => [...prev, `❌ Cloning failed: ${errorMessage}`]);
    }
  };

  return (
    <div className="p-6 space-y-6 max-w-4xl mx-auto">
      <Card>
        <CardContent className="p-6 flex justify-center">
          <AuthButtons />
        </CardContent>
      </Card>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <Card>
          <CardContent className="p-6 space-y-2">
            <label className="font-semibold">🔍 Source Site URL</label>
            <Input value={sourceUrl} onChange={(e) => setSourceUrl(e.target.value)} />
            <Button onClick={loadItemsFromSite} className="mt-2">Load Site Items</Button>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="p-6 space-y-2">
            <label className="font-semibold">🎯 Target Site URL</label>
            <Input value={targetUrl} onChange={(e) => setTargetUrl(e.target.value)} />
          </CardContent>
        </Card>
      </div>

      <Card>
        <CardContent className="p-6 space-y-4">
          <label className="font-semibold">✅ What to Clone?</label>
          <div className="grid grid-cols-2 gap-4">
            <div className="flex items-center space-x-2">
              <Checkbox checked={cloneLists} onCheckedChange={val => typeof val === 'boolean' && setCloneLists(val)} />
              <span>Lists</span>
            </div>
            <div className="flex items-center space-x-2">
              <Checkbox checked={cloneLibraries} onCheckedChange={val => typeof val === 'boolean' && setCloneLibraries(val)} />
              <span>Document Libraries</span>
            </div>
            <div className="flex items-center space-x-2">
              <Checkbox checked={clonePages} onCheckedChange={val => typeof val === 'boolean' && setClonePages(val)} />
              <span>Pages</span>
            </div>
            <div className="flex items-center space-x-2">
              <Checkbox checked={cloneSite} onCheckedChange={val => typeof val === 'boolean' && setCloneSite(val)} />
              <span>Full Site</span>
            </div>
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardContent className="p-6 space-y-2">
          <label className="font-semibold">📂 Available Items</label>
          <div className="flex items-center space-x-4 mb-2">
            <Button
              variant="outline"
              size="sm"
              onClick={() => setItems(items.map(item => ({ ...item, selected: true })))}
            >
              Select All
            </Button>
            <Button
              variant="outline"
              size="sm"
              onClick={() => setItems(items.map(item => ({ ...item, selected: false })))}
            >
              Deselect All
            </Button>
          </div>
          {items.map((item, index) => (
            <div key={index} className="flex items-center space-x-2">
              <Checkbox
                checked={item.selected}
                onCheckedChange={() => toggleItem(index)}
              />
              <span>{item.name} ({item.type})</span>
            </div>
          ))}
        </CardContent>
      </Card>

      <Card>
        <CardContent className="p-6 space-y-4">
          <div className="flex gap-2">
            <Button onClick={startCloning}>▶️ Start Cloning</Button>
            <Button 
              variant="outline" 
              onClick={() => setLogs([])}
            >
              🧹 Clear Logs
            </Button>
          </div>
          <div className="pt-2">
            <label className="font-semibold">🔧 Logs:</label>
            <div className="bg-gray-100 p-2 rounded text-sm h-96 overflow-y-scroll">
              {logs.map((log, i) => (
                <div key={i}>&gt; {log}</div>
              ))}
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}